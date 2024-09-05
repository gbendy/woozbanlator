import * as XLSX from 'xlsx';
import { readFile, writeFile } from 'node:fs/promises';
import readline from 'node:readline/promises';
import { ReadLineOptions } from 'node:readline';
import { compile } from './expression-eval';

const MATCHREGEXPS: unique symbol = Symbol();
const DATESTRING: unique symbol = Symbol();

const consoleColours = {
  Reset: '\x1b[0m',
  Bright: '\x1b[1m',
  Dim: '\x1b[2m',
  Underscore: '\x1b[4m',
  Blink: '\x1b[5m',
  Reverse: '\x1b[7m',
  Hidden: '\x1b[8m',

  FgBlack: '\x1b[30m',
  FgRed: '\x1b[31m',
  FgGreen: '\x1b[32m',
  FgYellow: '\x1b[33m',
  FgBlue: '\x1b[34m',
  FgMagenta: '\x1b[35m',
  FgCyan: '\x1b[36m',
  FgWhite: '\x1b[37m',
  FgGray: '\x1b[90m',

  BgBlack: '\x1b[40m',
  BgRed: '\x1b[41m',
  BgGreen: '\x1b[42m',
  BgYellow: '\x1b[43m',
  BgBlue: '\x1b[44m',
  BgMagenta: '\x1b[45m',
  BgCyan: '\x1b[46m',
  BgWhite: '\x1b[47m',
  BgGray: '\x1b[100m'
};

type CategoryDefinition = {
  name: string;
  matches: Array<string>;
  [MATCHREGEXPS]: Array<RegExp>;
  reimbursementFormula?: string;
}

type Config = {
  outputFilename: string;
  debitCategories: Record<string, CategoryDefinition>;
  creditCategories: Record<string, CategoryDefinition>;
}

type Transaction = {
  Date: number;
  Amount: number;
  Description: string;
  [DATESTRING]: string;
};

type ExtendedTransaction = Transaction & {
  Category: string;
  Reimbursed: string;
  metadata: {
    isDebit: boolean;
    isIntlTransactionFee: boolean;
  }
}

enum ReadlineType {
  general = 'general',
  category = 'category',
  regex = 'regex',
  formula = 'formula',
  intlTransaction = 'international'
}

type Context = {
  config: Config;
  configChanged: boolean;
  workbook: XLSX.WorkBook;
  debitsAdded: number;
  creditsAdded: number;
  existingTransactions: Set<string>;
  dateTransactions: Map<number, Set<ExtendedTransaction>>;
  readLine?: readline.Interface;
  readLineHistory: Record<ReadlineType, Array<string>>;
  intlTransactionFees: Set<ExtendedTransaction>;
}

const intlTransactionFee = /^INTNL TRANSACTION FEE$/;

function fileError(e: unknown) {
  if (e instanceof Error) {
    console.error(e.message);
  } else {
    console.error(e);
  }
}

function fatalFileError(e: unknown) {
  fileError(e);
  process.exit(1);
}

async function loadJson(filename: string) {
  const data = await readFile(filename, { encoding: 'utf8' });
  return JSON.parse(data);
}

async function saveConfig(filename: string, config: Config) {
  return writeFile(filename, JSON.stringify(config, null, 2), { encoding: 'utf8' });
}

function readWorkbook(filename: string, mustExist = true) {
  try {
    return XLSX.readFile(filename, { dateNF: 'dd/mm/yyyy', cellNF: true });
  } catch (e) {
    if (!mustExist && (e as { code: string }).code === 'ENOENT') {
      return undefined;
    }
    fatalFileError(e);
  }
}

function writeWorkbook(filename: string, workbook: XLSX.WorkBook) {
  try {
    XLSX.writeFileXLSX(workbook, filename);
    return true;
  } catch (e) {
    fileError(e);
  }
}

function createWorkbook() {
  const workbook = XLSX.utils.book_new();
  const debitSheet = XLSX.utils.aoa_to_sheet([[ 'Category', 'Date', 'Amount', 'Reimbursed', 'Paid', 'Description']], { });
  const creditSheet = XLSX.utils.aoa_to_sheet([[ 'Category', 'Date', 'Amount', 'Description']]);

  XLSX.utils.book_append_sheet(workbook, debitSheet, 'Debit');
  XLSX.utils.book_append_sheet(workbook, creditSheet, 'Credit');

  return workbook;
}

function prepareRegexps(categories: Record<string, CategoryDefinition>) {
  for (const categoryName in categories) {
    categories[categoryName][MATCHREGEXPS] = categories[categoryName].matches.map(match => new RegExp(match));
  }
}

function transactionId(transaction: Transaction) {
  return transaction.Date.toString() + ':' + transaction.Amount.toString() + ':' + transaction.Description;
}

function addToDateStore(transaction: ExtendedTransaction, context: Context) {
  // also add to per date store
  if (!context.dateTransactions.has(transaction.Date)) {
    context.dateTransactions.set(transaction.Date, new Set<ExtendedTransaction>());
  }
  context.dateTransactions.get(transaction.Date)?.add(transaction);
}
function cacheTransaction(transaction: ExtendedTransaction, context: Context) {
  context.existingTransactions.add(transactionId(transaction));
  addToDateStore(transaction, context);
}

function addExtendedDebit(workSheet: XLSX.WorkSheet, transaction: ExtendedTransaction, context: Context) {
  const range = workSheet['!ref'];
  const newRow = parseInt(range?.split(':')[1].substring(1) as string) + 1;

  const reimbursedValue = transaction.Reimbursed ? compile(transaction.Reimbursed)({ amount: Math.abs(transaction.Amount) }) : undefined;
  const reimbursedFormula = transaction.Reimbursed ? transaction.Reimbursed.replace(/amount/g, `ABS(C${newRow})`) : undefined;

  const paidValue = transaction.Amount + (reimbursedValue ?? 0);
  const paidFormula = `C${newRow}+D${newRow}`;

  const row = [
    {
      v: transaction.Category,
      t: 's'
    },
    {
      v: transaction.Date,
      t: 'n',
      z: 'm/d/yy'
    },
    {
      v: transaction.Amount,
      t: 'n',
      z: '"$"#,##0.00;[Red]"$"#,##0.00'
    },
    {
      v: reimbursedValue,
      f: reimbursedFormula,
      t: 'n',
      z: '"$"#,##0.00;[Red]"$"#,##0.00'
    },
    {
      v: paidValue,
      f: paidFormula,
      t: 'n',
      z: '"$"#,##0.00;[Red]"$"#,##0.00'
    },
    {
      v: transaction.Description,
      t: 's'
    }
  ];
  XLSX.utils.sheet_add_aoa(workSheet, [ row ], { origin: { r: -1, c: 0 } });
  context.debitsAdded++;
}

function addExtendedCredit(workSheet: XLSX.WorkSheet, transaction: ExtendedTransaction, context: Context) {
  const row = [
    {
      v: transaction.Category,
      t: 's'
    },
    {
      v: transaction.Date,
      t: 'n',
      z: 'm/d/yy'
    },
    {
      v: transaction.Amount,
      t: 'n',
      z: '"$"#,##0.00;[Red]"$"#,##0.00'
    },
    {
      v: transaction.Description,
      t: 's'
    }
  ];
  XLSX.utils.sheet_add_aoa(workSheet, [ row ], { origin: { r: -1, c: 0 } });
  context.creditsAdded++;
}

async function askYesNo(question: string, context: Context) {
  let answer:string = '';
  const rl = createReadline(context, ReadlineType.general);
  while (answer !== 'y' && answer !== 'n') {
    answer = await rl.question(`${question} (y/n)> `);
  }
  closeReadline(context);
  return answer === 'y';
}

async function processSigInt(context: Context) {
  closeReadline(context);
  if (context.debitsAdded || context.creditsAdded) {
    if (await askYesNo('Save workbook', context)) {
      writeWorkbook(context.config.outputFilename, context.workbook);
    }
    closeReadline(context);
  }
  if (context.configChanged) {
    if (await askYesNo('Save new categories', context)) {
      await saveConfig('config.json', context.config).catch(fatalFileError);
    }
  }
  process.exit(1);
};

function createReadline(context: Context, type: ReadlineType, options: Omit<ReadLineOptions, 'input'|'output'|'history'> = {}) {
  if (context.readLine) {
    context.readLine.close();
  }
  context.readLine = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    history: context.readLineHistory[type],
    ...options
  });
  context.readLine.on('SIGINT', processSigInt.bind(null, context));
  context.readLine.on('history', history => {
    context.readLineHistory[type] = history;
  });
  return context.readLine;
}

function closeReadline(context: Context) {
  context.readLine?.close();
  context.readLine = undefined;
}

function getPaddedAmount(transaction: Transaction) {
  return `${('$' + Math.abs(transaction.Amount).toFixed(2)).padStart(9, ' ')}`;
}

function fieldPrintTransaction(transaction: Transaction) {
  const amountColour = transaction.Amount < 0 ? `${consoleColours.FgRed}` : `${consoleColours.FgGreen}`;
  console.log(`${transaction[DATESTRING]} ${amountColour}${getPaddedAmount(transaction)}${consoleColours.Reset} ${transaction.Description}`);
}

function fieldPrintExtendedTransaction(transaction: ExtendedTransaction) {
  const amountColour = transaction.metadata.isDebit ? `${consoleColours.FgRed}` : `${consoleColours.FgGreen}`;
  console.log(`${transaction[DATESTRING]} ${amountColour}${getPaddedAmount(transaction)}${consoleColours.Reset} ${transaction.Category} ${transaction.Description}`);
}

async function getRegexForDescription(question: string, description: string, context: Context) {
  const regexRl = createReadline(context, ReadlineType.regex);
  let regex = await regexRl.question(`${question} >`);
  if (regex === '') {
    return;
  }
  while (!new RegExp(regex).test(description)) {
    console.error('Regex does not match description');
    regex = await regexRl.question(`${question} >`);
    if (regex === '') {
      return;
    }
  }
  return regex;
}

async function getReimbursementForTransaction(question: string, transaction: ExtendedTransaction, context: Context) {
  const reimburseRl = createReadline(context, ReadlineType.formula);
  let formula = await reimburseRl.question(`${question} >`);
  if (formula === '') {
    return;
  }
  let value = compile(formula)({ amount: Math.abs(transaction.Amount) });
  while (!await askYesNo(`Reimbursed $${value.toFixed(2)}`, context)) {
    formula = await createReadline(context, ReadlineType.formula).question(`${question} >`);
    if (formula === '') {
      return;
    }
    value = compile(formula)({ amount: Math.abs(transaction.Amount) });
  }
  return formula;
}

async function validateReimbursementForTransaction(transaction: ExtendedTransaction, formula: string, context: Context) {
  let value = compile(formula)({ amount: Math.abs(transaction.Amount) });
  while (!await askYesNo(`Reimbursed $${value.toFixed(2)}`, context)) {
    formula = await createReadline(context, ReadlineType.formula).question(`Reimbursement for ${transaction.Category} >`);
    if (formula === '') {
      return;
    }
    value = compile(formula)({ amount: Math.abs(transaction.Amount) });
  }
  transaction.Reimbursed = formula;
}

function isExtendedTransaction(transaction: Transaction | ExtendedTransaction): transaction is ExtendedTransaction {
  return !!(transaction as ExtendedTransaction).metadata;
}

async function categoriseTransaction(transaction: Transaction | ExtendedTransaction, context: Context) {
  const passedExtendedTransaction = isExtendedTransaction(transaction);
  const extendedTransaction: ExtendedTransaction = passedExtendedTransaction ? transaction : {
    ...transaction,
    Category: '',
    Reimbursed: '',
    metadata: {
      isDebit: transaction.Amount < 0,
      isIntlTransactionFee: intlTransactionFee.test(transaction.Description)
    }
  };

  if (!passedExtendedTransaction && extendedTransaction.metadata.isIntlTransactionFee) {
    return extendedTransaction;
  }

  const categories = extendedTransaction.metadata.isDebit ? context.config.debitCategories : context.config.creditCategories;

  // try and auto-assign a category
  const matches: Array<string> = [];

  if (passedExtendedTransaction) {
    // categorising international transactions so make sure that
    // transaction is one.
    if (!transaction.metadata.isIntlTransactionFee || transaction.Category) {
      throw 'Passed extended transaction with category is that is not an internation transaction fee';
    }
    const relatedTransactions = context.dateTransactions.get(transaction.Date);
    if (relatedTransactions) {
      const relatedCategories = [ ...new Set([...relatedTransactions.values()].map(t => t.Category)).values() ];
      const rl = createReadline(context, ReadlineType.intlTransaction, {
        completer: (line:string) => {
          const hits = relatedCategories.filter(category => category.startsWith(line));
          return [hits.length ? hits : relatedCategories, line];
        }
      });
      let input:string = undefined as unknown as string;
      while (input !== '' && relatedCategories.findIndex(match => match === input) < 0) {
        console.log('');
        fieldPrintTransaction(transaction);
        console.log('Transactions for same day as above international transaction fee');
        relatedTransactions.forEach(transaction => fieldPrintExtendedTransaction(transaction));
        input = await rl.question('Enter matching category or empty line to assign to new> ');
      }
      matches.length = 0;
      if (input !== '') {
        matches.push(input);
      }
    }
  } else {
    for (const categoryName in categories) {
      const category = categories[categoryName];
      for (const matchRegex of category[MATCHREGEXPS]) {
        if (matchRegex.test(transaction.Description)) {
          matches.push(categoryName);
        }
      }
    }
  }
  if (matches.length > 1) {
    // matched multiple categories, ask for which category to
    // allocate to or allow to enter a new category
    let input:string = undefined as unknown as string;
    const rl = createReadline(context, ReadlineType.category, {
      completer: (line:string) => {
        const hits = matches.filter(category => category.startsWith(line));
        return [hits.length ? hits : matches, line];
      }
    });
    while (input !== '' && matches.findIndex(match => match === input) < 0) {
      fieldPrintTransaction(transaction);
      console.log('Multiple categories matched:');
      matches.forEach(match => console.log(`  ${match}`));
      input = await rl.question('Enter matching category or empty line to create new category> ');
    }
    matches.length = 0;
    if (input !== '') {
      matches.push(input);
    }
  }
  if (matches.length === 1) {
    // A single category, assign it and apply reimbursement if defined
    extendedTransaction.Category = matches[0];
    fieldPrintExtendedTransaction(extendedTransaction);
    if (categories[extendedTransaction.Category]?.reimbursementFormula) {
      await validateReimbursementForTransaction(extendedTransaction, categories[extendedTransaction.Category].reimbursementFormula as string, context);
    }
  } else if (matches.length === 0) {
    // No match, ask for a category, and if it doesn't exist create it.
    fieldPrintTransaction(transaction);
    const existingCategories = Object.keys(categories);
    const rl = createReadline(context, ReadlineType.category, {
      completer: (line:string) => {
        const hits = existingCategories.filter(category => category.startsWith(line));
        return [hits.length ? hits : existingCategories, line];
      }
    });

    extendedTransaction.Category = await rl.question('Category for above transaction> ');
    if (!categories[extendedTransaction.Category]) {
      // Unknown category, create it if the user wants to.
      const regex = await getRegexForDescription(`Regex for ${extendedTransaction.Category}`, transaction.Description, context);
      context.configChanged = true;
      categories[extendedTransaction.Category] = {
        name: extendedTransaction.Category,
        matches: regex ? [ regex ] : [],
        [MATCHREGEXPS]: regex ? [ new RegExp(regex) ] : []
      };
      // assign reimbursement
      const formula = await getReimbursementForTransaction(`Reimbursement for ${extendedTransaction.Category}`, extendedTransaction, context);
      if (formula) {
        extendedTransaction.Reimbursed = formula;
        if (regex && await askYesNo(`Save formula to ${extendedTransaction.Category}`, context)) {
          categories[extendedTransaction.Category].reimbursementFormula = formula;
        }
      }
    } else {
      // existing category but it's regex doesn't match this description.
      // allow a new regex to be added and validate any forumula.
      const regex = await getRegexForDescription(`Regex to add to ${extendedTransaction.Category}`, transaction.Description, context);
      if (regex) {
        context.configChanged = true;
        categories[extendedTransaction.Category].matches.push(regex);
        categories[extendedTransaction.Category][MATCHREGEXPS].push(new RegExp(regex));
      }
      if (categories[extendedTransaction.Category].reimbursementFormula) {
        await validateReimbursementForTransaction(extendedTransaction, categories[extendedTransaction.Category].reimbursementFormula as string, context);
      }
    }
  }
  return extendedTransaction;
}

async function addTransaction(transaction: Transaction | ExtendedTransaction, context: Context) {
  const extendedTransaction = await categoriseTransaction(transaction, context);

  if (!(transaction as ExtendedTransaction).metadata?.isIntlTransactionFee && extendedTransaction.metadata.isIntlTransactionFee) {
    context.intlTransactionFees.add(extendedTransaction);
  } else {
    if (extendedTransaction.metadata.isDebit) {
      addExtendedDebit(context.workbook.Sheets['Debit'], extendedTransaction, context);
    } else {
      addExtendedCredit(context.workbook.Sheets['Credit'], extendedTransaction, context);
    }
    cacheTransaction(extendedTransaction, context);
  }
}

async function run(argv: Array<string>) {
  if (argv.length === 0) {
    console.error('No files given');
    process.exit(1);
  }
  const config = await loadJson('./config.json').catch(fatalFileError) as Config;

  if (!config.outputFilename) {
    fatalFileError('No outputFilename configured');
  }
  if (!config.creditCategories) {
    config.creditCategories = {};
  }
  if (!config.debitCategories) {
    config.debitCategories = {};
  }

  const outputWorkbook = readWorkbook(config.outputFilename, false) ?? createWorkbook();

  // parse all our input files first
  const files = argv.map(file => readWorkbook(file)) as Array<XLSX.WorkBook>;

  // prepare regexps for all categories
  prepareRegexps(config.debitCategories);
  prepareRegexps(config.creditCategories);

  const context: Context = {
    config,
    configChanged: false,
    workbook: outputWorkbook,
    debitsAdded: 0,
    creditsAdded: 0,
    existingTransactions: new Set<string>(),
    dateTransactions: new Map<number, Set<ExtendedTransaction>>(),
    readLineHistory: {} as Record<ReadlineType, []>,
    intlTransactionFees: new Set<ExtendedTransaction>
  };

  // make cache of existing entries
  for (const sheetName of [ 'Debit', 'Credit' ]) {
    const workSheet = outputWorkbook.Sheets[sheetName];
    const numRows = parseInt(workSheet['!ref']?.split(':')[1].substring(1) as string);

    for (let row=2; row<=numRows; row++) {
      const transaction: ExtendedTransaction = sheetName === 'Debit' ? {
        Category: workSheet[`A${row}`]?.w as string,
        Date: workSheet[`B${row}`].v as number,
        [DATESTRING]: workSheet[`B${row}`].w as string,
        Amount: workSheet[`C${row}`].v as number,
        Reimbursed: (workSheet[`D${row}`]?.f ?? '' as string).replace(new RegExp(`ABS\\(C${row}\\)`,'g'), 'amount'),
        Description: workSheet[`F${row}`].v as string,
        metadata: {
          isDebit: workSheet[`C${row}`].v < 0,
          isIntlTransactionFee: false
        }
      } : {
        Category: workSheet[`A${row}`]?.w as string,
        Date: workSheet[`B${row}`].v as number,
        [DATESTRING]: workSheet[`B${row}`].w as string,
        Amount: workSheet[`C${row}`].v as number,
        Reimbursed: '',
        Description: workSheet[`D${row}`].v as string,
        metadata: {
          isDebit: workSheet[`C${row}`].v < 0,
          isIntlTransactionFee: false
        }
      };
      cacheTransaction(transaction, context);
    }
  }

  process.on('SIGINT', processSigInt.bind(null, context));
  let currentDate: number | undefined;
  for (const workbook of files) {
    const workSheet = workbook.Sheets[workbook.SheetNames[0]];
    const numRows = parseInt(workSheet['!ref']?.split(':')[1].substring(1) as string);
    const sheetData: Array<Transaction> = [];

    for (let row=1; row<=numRows; row++) {
      sheetData.push({
        Date: workSheet[`A${row}`].v as number,
        [DATESTRING]: workSheet[`A${row}`].w as string,
        Amount: workSheet[`B${row}`].v as number,
        Description: workSheet[`C${row}`].v as string,
      });
    }

    sheetData.sort((a, b) => {
      if (a.Date < b.Date) {
        return -1;
      } else if (a.Date > b.Date) {
        return 1;
      }
      return 0;
    });
    if (!currentDate) {
      currentDate = sheetData[0].Date;
    }
    for (const transaction of sheetData) {
      if (transaction.Date !== currentDate) {
        // moved on to new day, check for international transaction fees to process
        if (context.intlTransactionFees.size) {
          for (const intlTransaction of context.intlTransactionFees) {
            await addTransaction(intlTransaction, context);
          }
          context.intlTransactionFees.clear();
        }
        currentDate = transaction.Date;
      }
      if (!context.existingTransactions.has(transactionId(transaction))) {
        await addTransaction(transaction, context);
      }
    }
  }
  if (context.debitsAdded || context.creditsAdded) {
    writeWorkbook(config.outputFilename, outputWorkbook);
  }
  if (context.configChanged) {
    await saveConfig('config.json', context.config).catch(fatalFileError);
  }
}

run(process.argv.splice(2));
