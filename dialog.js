/*
 * ATTENTION: The "eval" devtool has been used (maybe by default in mode: "development").
 * This devtool is neither made for production nor for readable output files.
 * It uses "eval()" calls to create a separate source file in the browser devtools.
 * If you are trying to read the output file, select a different devtool (https://webpack.js.org/configuration/devtool/)
 * or disable the default devtool with "devtool: false".
 * If you are looking for production-ready output files, see mode: "production" (https://webpack.js.org/configuration/mode/).
 */
/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/dialog/dialog.ts"
/*!******************************!*\
  !*** ./src/dialog/dialog.ts ***!
  \******************************/
() {

eval("{\nOffice.onReady(() => {\n    var _a, _b;\n    // The dialog receives folder info via the URL hash: #folderName=...&folderId=...\n    const params = new URLSearchParams(window.location.hash.slice(1));\n    const folderName = (_a = params.get(\"folderName\")) !== null && _a !== void 0 ? _a : \"Unknown folder\";\n    const folderId = (_b = params.get(\"folderId\")) !== null && _b !== void 0 ? _b : \"\";\n    document.getElementById(\"folderName\").textContent = folderName;\n    document.getElementById(\"move\").addEventListener(\"click\", () => {\n        Office.context.ui.messageParent(JSON.stringify({ action: \"move\", folderId }));\n    });\n    document.getElementById(\"skip\").addEventListener(\"click\", () => {\n        Office.context.ui.messageParent(JSON.stringify({ action: \"skip\" }));\n    });\n});\n\n\n//# sourceURL=webpack://outlook-filer/./src/dialog/dialog.ts?\n}");

/***/ }

/******/ 	});
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module can't be inlined because the eval devtool is used.
/******/ 	var __webpack_exports__ = {};
/******/ 	__webpack_modules__["./src/dialog/dialog.ts"]();
/******/ 	
/******/ })()
;