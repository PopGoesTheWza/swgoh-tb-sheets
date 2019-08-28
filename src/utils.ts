// tslint:disable: max-classes-per-file

namespace utils {
  export const EMOJI_KEYCAP_DIGITS = ['0️⃣', '1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣'];

  export function getActiveSheet() {
    return SPREADSHEET.getActiveSheet();
  }

  export function shrinkAllSheets() {
    SPREADSHEET.getSheets().forEach((sheet) => {
      const extraRows = 1;
      const extraColumns = 1;
      let { lastRow, lastColumn } = getSheetBoundaries(sheet);
      lastRow += extraRows;
      lastColumn += extraColumns;
      const maxRows = sheet.getMaxRows();
      const maxColumns = sheet.getMaxColumns();
      if (maxRows > lastRow) {
        sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      }
      if (maxColumns > lastColumn) {
        sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
      }
    });
  }

  function getSheetBoundaries(sheet: Spreadsheet.Sheet) {
    const dim = { lastColumn: 1, lastRow: 1 };
    sheet
      .getDataRange()
      .getMergedRanges()
      .forEach((e) => {
        const lastColumn = e.getLastColumn();
        const lastRow = e.getLastRow();
        if (lastColumn > dim.lastColumn) {
          dim.lastColumn = lastColumn;
        }
        if (lastRow > dim.lastRow) {
          dim.lastRow = lastRow;
        }
      });
    const rowCount = sheet.getMaxRows();
    const columnCount = sheet.getMaxColumns();
    const dataRange = sheet.getRange(1, 1, rowCount, columnCount).getValues();
    for (let rowIndex = rowCount; rowIndex > 0; rowIndex -= 1) {
      const row = dataRange[rowIndex - 1];
      if (row.join('').length > 0) {
        for (let columnIndex = columnCount; columnIndex > dim.lastColumn; columnIndex -= 1) {
          if (`${row[columnIndex - 1]}`.length > 0) {
            dim.lastColumn = columnIndex;
          }
        }
        if (dim.lastRow < rowIndex) {
          dim.lastRow = rowIndex;
        }
      }
    }
    return dim;
  }

  export function setActiveSheet(sheet: Spreadsheet.Sheet | string, failover?: Spreadsheet.Sheet | string) {
    const target = typeof sheet === 'string' ? SPREADSHEET.getSheetByName(sheet) : sheet;
    const current = SPREADSHEET.getActiveSheet();
    if (target !== current) {
      if (target && !target.isSheetHidden()) {
        target.activate();
      } else if (typeof failover === 'object' && failover.activate && typeof failover.activate === 'function') {
        failover.activate();
      } else if (typeof failover === 'string') {
        const someSheet = SPREADSHEET.getSheetByName(failover);
        if (someSheet) {
          someSheet.activate();
        } // TODO: else
      }
    }
    // return SPREADSHEET.getActiveSheet();
  }

  export function getSheetByNameOrDie(name: string) {
    const sheet = SPREADSHEET.getSheetByName(name);
    if (sheet) {
      return sheet;
    } else {
      throw new Error(`Missing sheet labeled '${name}'`);
    }
  }

  /** string[] callback for case insensitive alphabetical sort */
  export function caseInsensitive(a: string, b: string): number {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  }

  export function clone<T>(mutable: T): T {
    return JSON.parse(JSON.stringify(mutable));
  }

  class Queue<T> {
    private store: T[] = [];

    public push(value: T) {
      this.store.push(value);
    }

    public pop(): T | undefined {
      return this.store.shift();
    }

    public isEmpty() {
      return this.store.length === 0;
    }
  }

  export type SpooledTask = () => void;

  export class Spooler {
    private queue: Queue<SpooledTask>;

    constructor() {
      this.queue = new Queue();
    }

    public attach(range: Spreadsheet.Range): SpooledRange {
      return new SpooledRange(this, range);
    }

    public add(task: SpooledTask) {
      this.queue.push(task);
    }

    public commit() {
      while (!this.queue.isEmpty()) {
        this.queue.pop()!();
      }
    }
  }

  export class SpooledRange {
    private readonly range: Spreadsheet.Range;
    private readonly spooler: Spooler;

    constructor(spooler: Spooler, range: Spreadsheet.Range) {
      this.range = range;
      this.spooler = spooler;
    }
    public clear(options?: object) {
      let immutable: any;
      if (typeof options === 'object') {
        immutable = { ...options };
      }

      const range = this.range;
      this.addTask(() => range.clear(immutable));

      return this;
    }

    public clearContent() {
      const range = this.range;
      this.addTask(() => range.clearContent());

      return this;
    }

    public clearDataValidations() {
      const range = this.range;
      this.addTask(() => range.clearDataValidations());

      return this;
    }

    public clearFormat() {
      const range = this.range;
      this.addTask(() => range.clearFormat());

      return this;
    }

    public clearNote() {
      const range = this.range;
      this.addTask(() => range.clearNote());

      return this;
    }

    public offset(rowOffset: number, columnOffset: number, numRows?: number, numColumn?: number) {
      const range = this.range;
      const offset = range.offset(
        rowOffset,
        columnOffset,
        numRows || range.getNumRows(),
        numColumn || range.getNumColumns(),
      );

      return new SpooledRange(this.spooler, offset);
    }

    public setBackground(color: string) {
      const range = this.range;
      this.addTask(() => range.setBackground(color));

      return this;
    }

    public setBackgroundRGB(red: number, green: number, blue: number) {
      const range = this.range;
      this.addTask(() => range.setBackgroundRGB(red, green, blue));

      return this;
    }

    public setBackgrounds(colors: string[][]) {
      const range = this.range;
      const immutable = clone(colors);
      this.addTask(() => range.setBackgrounds(immutable));

      return this;
    }

    public setBorder(
      top: boolean,
      left: boolean,
      bottom: boolean,
      right: boolean,
      vertical: boolean,
      horizontal: boolean,
      color: string, // or null
      style: Spreadsheet.BorderStyle, // or null
    ) {
      const range = this.range;
      this.addTask(() => range.setBorder(top, left, bottom, right, vertical, horizontal, color, style));

      return this;
    }

    public setDataValidation(rule: Spreadsheet.DataValidation) {
      const range = this.range;
      this.addTask(() => range.setDataValidation(rule));

      return this;
    }

    public setDataValidations(rules: Spreadsheet.DataValidation[][]) {
      const range = this.range;
      this.addTask(() => range.setDataValidations(rules));

      return this;
    }

    public setFontColor(color: string) {
      const range = this.range;
      this.addTask(() => range.setFontColor(color));

      return this;
    }

    public setFontColors(colors: string[][]) {
      const range = this.range;
      const immutable = clone(colors);
      this.addTask(() => range.setFontColors(immutable));

      return this;
    }

    public setFontFamilies(fontFamilies: string[][]) {
      const range = this.range;
      const immutable = clone(fontFamilies);
      this.addTask(() => range.setFontFamilies(immutable));

      return this;
    }

    public setFontFamily(fontFamily: string) {
      const range = this.range;
      this.addTask(() => range.setFontFamily(fontFamily));

      return this;
    }

    public setFontLine(fontLine: 'underline' | 'line-through' | 'none') {
      const range = this.range;
      this.addTask(() => range.setFontLine(fontLine));

      return this;
    }

    public setFontLines(fontLines: Array<Array<'underline' | 'line-through' | 'none'>>) {
      const range = this.range;
      const immutable = clone(fontLines);
      this.addTask(() => range.setFontLines(immutable));

      return this;
    }

    public setFontSize(fontSize: number) {
      const range = this.range;
      this.addTask(() => range.setFontSize(fontSize));

      return this;
    }

    public setFontSizes(fontSizes: number[][]) {
      const range = this.range;
      const immutable = clone(fontSizes);
      this.addTask(() => range.setFontSizes(immutable));

      return this;
    }

    public setFontStyle(fontStyle: 'italic' | 'normal') {
      const range = this.range;
      this.addTask(() => range.setFontStyle(fontStyle));

      return this;
    }

    public setFontStyles(fontStyles: Array<Array<'italic' | 'normal'>>) {
      const range = this.range;
      const immutable = clone(fontStyles);
      this.addTask(() => range.setFontStyles(immutable));

      return this;
    }

    public setFontWeight(fontWeight: 'bold' | 'normal') {
      const range = this.range;
      this.addTask(() => range.setFontWeight(fontWeight));

      return this;
    }

    public setFontWeights(fontWeights: Array<Array<'bold' | 'normal'>>) {
      const range = this.range;
      const immutable = clone(fontWeights);
      this.addTask(() => range.setFontWeights(immutable));

      return this;
    }

    public setFormula(formula: string) {
      const range = this.range;
      this.addTask(() => range.setFormula(formula));

      return this;
    }

    public setFormulaR1C1(formula: string) {
      const range = this.range;
      this.addTask(() => range.setFormulaR1C1(formula));

      return this;
    }

    public setFormulas(formulas: string[][]) {
      const range = this.range;
      const immutable = clone(formulas);
      this.addTask(() => range.setFormulas(immutable));

      return this;
    }

    public setFormulasR1C1(formulas: string[][]) {
      const range = this.range;
      const immutable = clone(formulas);
      this.addTask(() => range.setFormulasR1C1(immutable));

      return this;
    }

    public setHorizontalAlignment(alignment: 'left' | 'center' | 'right') {
      const range = this.range;
      this.addTask(() => range.setHorizontalAlignment(alignment));

      return this;
    }

    public setHorizontalAlignments(alignments: Array<Array<'left' | 'center' | 'right'>>) {
      const range = this.range;
      const immutable = clone(alignments);
      this.addTask(() => range.setHorizontalAlignments(immutable));

      return this;
    }

    public setNote(note: string) {
      const range = this.range;
      this.addTask(() => range.setNote(note));

      return this;
    }

    public setNotes(notes: string[][]) {
      const range = this.range;
      const immutable = clone(notes);
      this.addTask(() => range.setNotes(immutable));

      return this;
    }

    public setNumberFormat(numberFormat: string) {
      const range = this.range;
      this.addTask(() => range.setNumberFormat(numberFormat));

      return this;
    }

    public setNumberFormats(numberFormats: string[][]) {
      const range = this.range;
      const immutable = clone(numberFormats);
      this.addTask(() => range.setNumberFormats(immutable));

      return this;
    }

    public setShowHyperlink(showHyperlink: boolean) {
      const range = this.range;
      this.addTask(() => range.setShowHyperlink(showHyperlink));

      return this;
    }

    public setTextDirection(textDirection: Spreadsheet.TextDirection) {
      const range = this.range;
      this.addTask(() => range.setTextDirection(textDirection));

      return this;
    }

    public setTextDirections(textDirections: Spreadsheet.TextDirection[][]) {
      const range = this.range;
      const immutable = clone(textDirections);
      this.addTask(() => range.setTextDirections(immutable));

      return this;
    }

    public setTextRotation(rotation: Spreadsheet.TextRotation) {
      const range = this.range;
      this.addTask(() => range.setTextRotation(rotation));

      return this;
    }

    public setTextRotations(rotations: Spreadsheet.TextRotation[][]) {
      const range = this.range;
      const immutable = clone(rotations);
      this.addTask(() => range.setTextRotations(immutable));

      return this;
    }

    public setValue(value: any) {
      const range = this.range;
      this.addTask(() => range.setValue(value));

      return this;
    }

    public setValues(values: any[][]) {
      const range = this.range;
      const immutable = clone(values);
      this.addTask(() => range.setValues(immutable));

      return this;
    }

    public setVerticalAlignment(alignment: 'top' | 'middle' | 'bottom') {
      const range = this.range;
      this.addTask(() => range.setVerticalAlignment(alignment));

      return this;
    }

    public setVerticalAlignments(alignments: Array<Array<'top' | 'middle' | 'bottom'>>) {
      const range = this.range;
      const immutable = clone(alignments);
      this.addTask(() => range.setVerticalAlignments(immutable));

      return this;
    }

    public setVerticalText(isVertical: boolean) {
      const range = this.range;
      this.addTask(() => range.setVerticalText(isVertical));

      return this;
    }

    public setWrap(isWrapEnabled: boolean) {
      const range = this.range;
      this.addTask(() => range.setWrap(isWrapEnabled));

      return this;
    }

    public setWrapStrategies(strategies: Spreadsheet.WrapStrategy[][]) {
      const range = this.range;
      const immutable = clone(strategies);
      this.addTask(() => range.setWrapStrategies(immutable));

      return this;
    }

    public setWrapStrategy(strategy: Spreadsheet.WrapStrategy) {
      const range = this.range;
      this.addTask(() => range.setWrapStrategy(strategy));

      return this;
    }

    public setWraps(isWrapEnabled: boolean[][]) {
      const range = this.range;
      const immutable = clone(isWrapEnabled);
      this.addTask(() => range.setWraps(immutable));

      return this;
    }

    protected addTask(task: SpooledTask) {
      this.spooler.add(task);
    }
  }
}
