class Cell {
  constructor(value = '', formula = null) {
    this.value = value;
    this.formula = formula;
    this.style = {};
    this.isMerged = false;
    this.mergeInfo = null; // { startRow, startCol, endRow, endCol }
  }

  setValue(value) {
    this.value = value;
    return this;
  }

  getValue() {
    return this.value;
  }

  setFormula(formula) {
    this.formula = formula;
    return this;
  }

  getFormula() {
    return this.formula;
  }

  setStyle(styleObj) {
    this.style = { ...this.style, ...styleObj };
    return this;
  }

  getStyle() {
    return this.style;
  }

  setMerged(mergeInfo) {
    this.isMerged = true;
    this.mergeInfo = mergeInfo;
    return this;
  }

  clearMerge() {
    this.isMerged = false;
    this.mergeInfo = null;
    return this;
  }

  clone() {
    const cell = new Cell(this.value, this.formula);
    cell.style = { ...this.style };
    cell.isMerged = this.isMerged;
    cell.mergeInfo = this.mergeInfo ? { ...this.mergeInfo } : null;
    return cell;
  }

  toJSON() {
    return {
      value: this.value,
      formula: this.formula,
      style: this.style,
      isMerged: this.isMerged,
      mergeInfo: this.mergeInfo
    };
  }
}

module.exports = Cell;
