import * as XLSX from 'xlsx';
import { readFile } from 'node:fs/promises';
import { compile } from './expression-eval';

type CategoryDefinition = {
  name: string;
  matches: Array<string>;
  reimbursementFormula?: string;
}

type Config = {
  outputFilename: string;
  debitCategories: Record<string, CategoryDefinition>;
  creditCategories: Record<string, CategoryDefinition>;
}

const TransactionHeader = [ 'Date', 'Amount', 'Description', 'Description2' ];
type Transaction = {
  Date: number;
  Amount: number;
  Description: string;
};

type ExtendedTransaction = Transaction & {
  Category: string;
  Reimbursed: string;
}

type Statistics = {
  debitsAdded: number;
  creditsAdded: number;
}

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

function readWorkbook(filename: string, mustExist = true) {
  try {
    return XLSX.readFile(filename, { dateNF: 'dd/mm/yyyy', cellNF: true, });
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
  const debitSheet = XLSX.utils.aoa_to_sheet([[ 'Category', 'Date', 'Amount', 'Reimbursed', 'Description']], { });
  const creditSheet = XLSX.utils.aoa_to_sheet([[ 'Category', 'Date', 'Amount', 'Reimbursed', 'Description']]);

  XLSX.utils.book_append_sheet(workbook, debitSheet, 'Debit');
  XLSX.utils.book_append_sheet(workbook, creditSheet, 'Credit');

  return workbook;
}

function transactionId(transaction: Transaction) {
  return transaction.Date.toString() + ':' + transaction.Amount.toString() + ':' + transaction.Description;
}

function addExtendedTransaction(workSheet: XLSX.WorkSheet, transaction: ExtendedTransaction) {
  const range = workSheet['!ref'];
  const newRow = parseInt(range?.split(':')[1][1] as string) + 1;

  const reimbursedAmount = compile(transaction.Reimbursed)({ amount: transaction.Amount });
  const reimbursedForumula = transaction.Reimbursed.replace(/amount/g, `C${newRow}`);
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
      v: reimbursedAmount,
      f: reimbursedForumula,
      t: 'n',
      z: '"$"#,##0.00;[Red]"$"#,##0.00'
    },
    {
      v: transaction.Description,
      t: 's'
    }
  ];
  XLSX.utils.sheet_add_aoa(workSheet, [ row ], { origin: { r: -1, c: 0 } });
}
function processTransaction(workbook: XLSX.WorkBook, transaction: Transaction, statistics: Statistics) {
  const extendedTransaction = { ...transaction } as ExtendedTransaction;
  extendedTransaction.Category = 'Stuff';
  extendedTransaction.Reimbursed = 'amount';

  const workSheet = transaction.Amount < 0 ? workbook.Sheets['Debit'] : workbook.Sheets['Credit'];
  addExtendedTransaction(workSheet, extendedTransaction);

  if (transaction.Amount < 0) {
    statistics.debitsAdded++;
  } else {
    statistics.creditsAdded++;
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

  // make cache of existing entries
  const existingTransactions = new Set<string>();
  const dateTransactions = new Map<number, Set<ExtendedTransaction>>();
  for (const sheetName of [ 'Debit', 'Credit' ]) {
    const sheetData = XLSX.utils.sheet_to_json(outputWorkbook.Sheets[sheetName], {
    }) as Array<ExtendedTransaction>;
    for (const transaction of sheetData) {
      existingTransactions.add(transactionId(transaction));
      // also add to per date store
      if (!dateTransactions.has(transaction.Date)) {
        dateTransactions.set(transaction.Date, new Set<ExtendedTransaction>());
      }
      dateTransactions.get(transaction.Date)?.add(transaction);
    }

  }
  const statistics = {
    debitsAdded: 0,
    creditsAdded: 0
  };

  for (const workbook of files) {
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
      header: TransactionHeader
    }) as Array<Transaction>;

    sheetData.sort((a, b) => {
      if (a.Date < b.Date) {
        return -1;
      } else if (a.Date > b.Date) {
        return 1;
      }
      return 0;
    });
    for (const transaction of sheetData) {
      if (!existingTransactions.has(transactionId(transaction))) {
        processTransaction(outputWorkbook, transaction, statistics);
      }
    }
  }
  if (statistics.debitsAdded || statistics.creditsAdded) {
    writeWorkbook(config.outputFilename, outputWorkbook);
  }
}

run(process.argv.splice(2));
