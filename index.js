const ROW_COUNT = 20;
const COLUMN_COUNT = 6;

let graph;
let parser;

// Reference to cell of currently executing formula.
let currentCellId = null;

function initTable() {
  for (let i = 0; i < ROW_COUNT + 1; i++) {
    let row = document.getElementById("table").insertRow(-1);

    for (let j = 0; j < COLUMN_COUNT + 1; j++) {
      const letter = String.fromCharCode("A".charCodeAt(0) + j - 1);
      const id = letter + i;

      if (i && j) {
        const input = createInput(id);
        row.insertCell(-1).appendChild(input);
        graph.setNode(id, { value: "", formula: undefined });
      } else {
        row.insertCell(-1).innerHTML = i || letter;
      }
    }
  }
}

function createInput(id) {
  const input = document.createElement("input");
  input.id = id;
  input.onblur = onInputBlur;
  input.onfocus = onInputFocus;
  return input;
}

function onInputBlur(event) {
  const id = event.target.id;
  const value = event.target.value;

  const properties = {
    value: undefined,
    formula: undefined
  };

  if (value.startsWith("=")) {
    properties.formula = value;
  } else {
    properties.value = value;
  }

  graph.setNode(id, properties);
  updateCell(id);
}

function onInputFocus(event) {
  const id = event.target.id;
  const element = document.getElementById(id);
  const data = graph.node(id);

  element.value = data.formula || data.value;
}

/**
 * Update a cell value.
 * if its a formula parse and execute formula.
 */
function updateCell(id) {
  const element = document.getElementById(id);
  const node = graph.node(id);

  if (node.formula) {
    const { result, error } = executeFormula(id, node.formula);
    element.value = error || result;
  } else {
    element.value = node.value;
  }

  graph.setNode(id, { ...node, value: element.value });
  // recursively follows all of the edges. This is not as efficient as
  // depth first search since this will follow all the routes blindly
  // it should be good enough for this case.

  // beware of cyclic dependencies. Need to give a warning here
  graph.outEdges(id).forEach(e => e.w !== id && updateCell(e.w));
}

/**
 * Execute formula with setting the current cell.
 * Everytime we request a cell during formula execution
 * we will build the dependency graph.
 */
function executeFormula(id, formula) {
  currentCellId = id;
  const result = parser.parse(formula.slice(1));
  currentCellId = null;

  return result;
}

/**
 * Cell value getter listener.
 * Resolves cell value and builds dependency graph.
 */
function onCallCellValue(coordinate, done) {
  const id = `${coordinate.column.label}${coordinate.row.label}`;
  graph.setEdge(id, currentCellId);

  const { value } = graph.node(id);
  done(value);
}

/**
 * Cell range value getter listener.
 * Resolves cell value and builds dependency graph.
 */
function onCallRangeValue(startcoordinate, endcoordinate, done) {
  const fragment = [];

  for (
    let row = startcoordinate.row.index;
    row <= endcoordinate.row.index;
    row++
  ) {
    const columnFragment = [];

    for (
      let col = startcoordinate.column.index;
      col <= endcoordinate.column.index;
      col++
    ) {
      const id = `${String.fromCharCode(65 + col)}${row + 1}`;
      graph.setEdge(id, currentCellId);

      const { value } = graph.node(id);
      columnFragment.push(value);
    }
    fragment.push(columnFragment);
  }

  done(fragment);
}

function main() {
  graph = new graphlib.Graph();
  parser = new formulaParser.Parser();

  parser.on("callCellValue", onCallCellValue);
  parser.on("callRangeValue", onCallRangeValue);

  // For exploration
  window.graph = graph;
  window.parser = parser;

  initTable();
}
