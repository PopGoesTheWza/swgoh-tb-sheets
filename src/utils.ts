
namespace utils {

  /** string[] callback for case insensitive alphabetical sort */
  export function caseInsensitive(a: string, b: string): number {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  }

  export function clone<T>(mutable: T): T {
    return JSON.parse(JSON.stringify(mutable));
  }

  class Queue<T> {

    private store: T[] = [];

    push(value: T) {
      this.store.push(value);
    }

    pop(): T | undefined {
      return this.store.shift();
    }

    isEmpty() {
      return this.store.length === 0;
    }

  }

  export type SpooledTask = () => void;

  export class Spooler {

    private queue: Queue<SpooledTask>;

    constructor() {
      this.queue = new Queue();
    }

    attach(range: Spreadsheet.Range): SpooledRange {
      return new SpooledRange(this, range);
    }

    add(task: SpooledTask) {
      this.queue.push(task);
    }

    commit() {
      while (!this.queue.isEmpty()) {
        this.queue.pop()();
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

    protected addTask(task: SpooledTask) {
      this.spooler.add(task);
    }

    clear(options: Object|undefined) {

      let immutable = undefined;
      if (typeof options === 'object') {
        immutable = Object.assign({}, options);
      }

      const range = this.range;
      this.addTask(() => range.clear(immutable));

      return this;
    }

    clearContent() {
      const range = this.range;
      this.addTask(() => range.clearContent());

      return this;
    }

    clearDataValidations() {
      const range = this.range;
      this.addTask(() => range.clearDataValidations());

      return this;
    }

    clearFormat(options: Object|undefined) {
      const range = this.range;
      this.addTask(() => range.clearFormat());

      return this;
    }

    clearNote(options: Object|undefined) {
      const range = this.range;
      this.addTask(() => range.clearNote());

      return this;
    }

    offset(rowOffset: number, columnOffset: number, numRows?: number, numColumn?: number) {

      const range = this.range;
      const offset = range.offset(
        rowOffset,
        columnOffset,
        numRows || range.getNumRows(),
        numColumn || range.getNumColumns(),
      );

      return new SpooledRange(this.spooler, offset);
    }

    setBackground(color: string) {
      const range = this.range;
      this.addTask(() => range.setBackground(color));

      return this;
    }

    setBackgroundRGB(red: number, green: number, blue: number) {
      const range = this.range;
      this.addTask(() => range.setBackgroundRGB(red, green, blue));

      return this;
    }

    setBackgrounds(colors: string[][]) {
      const range = this.range;
      const immutable = clone(colors);
      this.addTask(() => range.setBackgrounds(immutable));

      return this;
    }

    setBorder(
      top: boolean,
      left: boolean,
      bottom: boolean,
      right: boolean,
      vertical: boolean,
      horizontal: boolean,
      color?: string,
      style?: Spreadsheet.BorderStyle,
    ) {
      const range = this.range;
      this.addTask(() => range.setBorder(
        top,
        left,
        bottom,
        right,
        vertical,
        horizontal,
        color,
        style,
      ));

      return this;
    }

    setDataValidation(rule: Spreadsheet.DataValidation) {
      const range = this.range;
      this.addTask(() => range.setDataValidation(rule));

      return this;
    }

    setDataValidations(rules: Spreadsheet.DataValidation[][]) {
      const range = this.range;
      this.addTask(() => range.setDataValidations(rules));

      return this;
    }

    setFontColor(color: string) {
      const range = this.range;
      this.addTask(() => range.setFontColor(color));

      return this;
    }

    setFontColors(colors: string[][]) {
      const range = this.range;
      const immutable = clone(colors);
      this.addTask(() => range.setFontColors(immutable));

      return this;
    }

    setFontFamilies(fontFamilies: string[][]) {
      const range = this.range;
      const immutable = clone(fontFamilies);
      this.addTask(() => range.setFontFamilies(immutable));

      return this;
    }

    setFontFamily(fontFamily: string) {
      const range = this.range;
      this.addTask(() => range.setFontFamily(fontFamily));

      return this;
    }

    setFontLine(fontLine: 'underline'|'line-through'|'none') {
      const range = this.range;
      this.addTask(() => range.setFontLine(fontLine));

      return this;
    }

    setFontLines(fontLines: ('underline'|'line-through'|'none')[][]) {
      const range = this.range;
      const immutable = clone(fontLines);
      this.addTask(() => range.setFontLines(immutable));

      return this;
    }

    setFontSize(fontSize: number) {
      const range = this.range;
      this.addTask(() => range.setFontSize(fontSize));

      return this;
    }

    setFontSizes(fontSizes: number[][]) {
      const range = this.range;
      const immutable = clone(fontSizes);
      this.addTask(() => range.setFontSizes(immutable));

      return this;
    }

    setFontStyle(fontStyle: 'italic'|'normal') {
      const range = this.range;
      this.addTask(() => range.setFontStyle(fontStyle));

      return this;
    }

    setFontStyles(fontStyles: ('italic'|'normal')[][]) {
      const range = this.range;
      const immutable = clone(fontStyles);
      this.addTask(() => range.setFontStyles(immutable));

      return this;
    }

    setFontWeight(fontWeight: 'bold'|'normal') {
      const range = this.range;
      this.addTask(() => range.setFontWeight(fontWeight));

      return this;
    }

    setFontWeights(fontWeights: ('bold'|'normal')[][]) {
      const range = this.range;
      const immutable = clone(fontWeights);
      this.addTask(() => range.setFontWeights(immutable));

      return this;
    }

    setFormula(formula: string) {
      const range = this.range;
      this.addTask(() => range.setFormula(formula));

      return this;
    }

    setFormulaR1C1(formula: string) {
      const range = this.range;
      this.addTask(() => range.setFormulaR1C1(formula));

      return this;
    }

    setFormulas(formulas: string[][]) {
      const range = this.range;
      const immutable = clone(formulas);
      this.addTask(() => range.setFormulas(immutable));

      return this;
    }

    setFormulasR1C1(formulas: string[][]) {
      const range = this.range;
      const immutable = clone(formulas);
      this.addTask(() => range.setFormulasR1C1(immutable));

      return this;
    }

    setHorizontalAlignment(alignment: 'left'|'center'|'right') {
      const range = this.range;
      this.addTask(() => range.setHorizontalAlignment(alignment));

      return this;
    }

    setHorizontalAlignments(alignments: ('left'|'center'|'right')[][]) {
      const range = this.range;
      const immutable = clone(alignments);
      this.addTask(() => range.setHorizontalAlignments(immutable));

      return this;
    }

    setNote(note: string) {
      const range = this.range;
      this.addTask(() => range.setNote(note));

      return this;
    }

    setNotes(notes: string[][]) {
      const range = this.range;
      const immutable = clone(notes);
      this.addTask(() => range.setNotes(immutable));

      return this;
    }

    setNumberFormat(numberFormat: string) {
      const range = this.range;
      this.addTask(() => range.setNumberFormat(numberFormat));

      return this;
    }

    setNumberFormats(numberFormats: string[][]) {
      const range = this.range;
      const immutable = clone(numberFormats);
      this.addTask(() => range.setNumberFormats(immutable));

      return this;
    }

    setShowHyperlink(showHyperlink: boolean) {
      const range = this.range;
      this.addTask(() => range.setShowHyperlink(showHyperlink));

      return this;
    }

    setTextDirection(textDirection: Spreadsheet.TextDirection) {
      const range = this.range;
      this.addTask(() => range.setTextDirection(textDirection));

      return this;
    }

    setTextDirections(textDirections: Spreadsheet.TextDirection[][]) {
      const range = this.range;
      const immutable = clone(textDirections);
      this.addTask(() => range.setTextDirections(immutable));

      return this;
    }

    setTextRotation(rotation: Spreadsheet.TextRotation) {
      const range = this.range;
      this.addTask(() => range.setTextRotation(rotation));

      return this;
    }

    setTextRotations(rotations: Spreadsheet.TextRotation[][]) {
      const range = this.range;
      const immutable = clone(rotations);
      this.addTask(() => range.setTextRotations(immutable));

      return this;
    }

    setValue(value: any) {
      const range = this.range;
      this.addTask(() => range.setValue(value));

      return this;
    }

    setValues(values: any[][]) {
      const range = this.range;
      const immutable = clone(values);
      this.addTask(() => range.setValues(immutable));

      return this;
    }

    setVerticalAlignment(alignment: 'top'|'middle'|'bottom') {
      const range = this.range;
      this.addTask(() => range.setVerticalAlignment(alignment));

      return this;
    }

    setVerticalAlignments(alignments: ('top'|'middle'|'bottom')[][]) {
      const range = this.range;
      const immutable = clone(alignments);
      this.addTask(() => range.setVerticalAlignments(immutable));

      return this;
    }

    setVerticalText(isVertical: boolean) {
      const range = this.range;
      this.addTask(() => range.setVerticalText(isVertical));

      return this;
    }

    setWrap(isWrapEnabled: boolean) {
      const range = this.range;
      this.addTask(() => range.setWrap(isWrapEnabled));

      return this;
    }

    setWrapStrategies(strategies: Spreadsheet.WrapStrategy[][]) {
      const range = this.range;
      const immutable = clone(strategies);
      this.addTask(() => range.setWrapStrategies(immutable));

      return this;
    }

    setWrapStrategy(strategy: Spreadsheet.WrapStrategy) {
      const range = this.range;
      this.addTask(() => range.setWrapStrategy(strategy));

      return this;
    }

    setWraps(isWrapEnabled: boolean[][]) {
      const range = this.range;
      const immutable = clone(isWrapEnabled);
      this.addTask(() => range.setWraps(immutable));

      return this;
    }

  }

}
