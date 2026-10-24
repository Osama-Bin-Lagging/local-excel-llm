// Offline Excel formula engine for the add-in and the test harness:
// fast-formula-parser parses and resolves references; formula.js supplies functions it lacks.
const FormulaParser = require("fast-formula-parser");
const formulajs = require("@formulajs/formulajs");
globalThis.FormulaParser = FormulaParser;
globalThis.formulajs = formulajs;
