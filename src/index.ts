/* eslint-disable @typescript-eslint/no-unused-vars */
import { uberLogger } from "./lib/uberLogger";
import { RSI_legacy, beta_legacy, hourlyRefresh } from "./legacy";
import { CellValue } from "./domain/@types";

const onOpen = () => {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("SCRIPTS")
    .addItem("🔁 example", "runExampleMenuItem")
    .addToUi();
};

const runHourlyRefresh = () => {
  uberLogger.init({ tabName: "LOGS" });
  try {
    hourlyRefresh();
  } catch (error) {
    uberLogger.error((error as Error).toString());
  }
};

const RSI = (
  symbolList: CellValue[],
  dataSymbol: CellValue[],
  dataDate: CellValue[],
  dataPrice: CellValue[]
) => {
  uberLogger.init({ tabName: "LOGS" });
  try {
    RSI_legacy(symbolList, dataSymbol, dataDate, dataPrice);
  } catch (error) {
    uberLogger.error((error as Error).toString());
  }
};

const beta = (
  symbolList: CellValue[],
  indices: CellValue[],
  dataSymbol: CellValue[],
  dataDate: CellValue[],
  dataPrice: CellValue[]
) => {
  uberLogger.init({ tabName: "LOGS" });
  try {
    beta_legacy(symbolList, indices, dataSymbol, dataDate, dataPrice);
  } catch (error) {
    uberLogger.error((error as Error).toString());
  }
};
