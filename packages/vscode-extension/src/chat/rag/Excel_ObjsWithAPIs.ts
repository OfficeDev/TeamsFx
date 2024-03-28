// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export const excelJsApiDocs = [
  {
    objName: "Word.Body",
    apiList: [
      {
        name: "Word.Body.getComments",
        description: "Get all the comments in the document body.",
        kind: "Method",
        signature: "Word.Body.getComments(): Word.CommentCollection",
        examples: [
          "const comments = context.document.body.getComments(); \n   comments.load(); \n  await context.sync();",
        ],
      },
      {
        name: "Word.Body.getHtml",
        description: "Gets an HTML representation of the body object. ",
        kind: "Method",
        signature: "Word.Body.getHtml(): OfficeExtension.ClientResult<string>",
        examples: [],
      },
    ],
  },
  {
    objName: "Word.Range",
    apiList: [
      {
        name: "Word.Range.getComments",
        description: "Get all the comments in the range or selection.",
        kind: "Method",
        signature: "Word.Range.getComments(): Word.CommentCollection",
        examples: [
          "const comments = context.document.getSelection().getComments(); \n   comments.load(); \n  await context.sync();",
        ],
      },
      {
        name: "Word.Range.getHtml",
        description: "Gets an HTML representation of the range object or current selection. ",
        kind: "Method",
        signature: "Word.Range.getHtml(): OfficeExtension.ClientResult<string>",
        examples: [],
      },
    ],
  },
  {
    objName: "Word.Paragraph",
    apiList: [
      {
        name: "Word.Paragraph.getComments",
        description: "Get all the comments in the paragraph.",
        kind: "Method",
        signature: "Word.Paragraph.getComments(): Word.CommentCollection",
        examples: [
          "const comments = context.document.paragraphs.getFirst().getComments(); \n   comments.load(); \n  await context.sync();",
        ],
      },
      {
        name: "Word.Paragraph.getHtml",
        description: "Gets an HTML representation of the paragraph.",
        kind: "Method",
        signature: "Word.Paragraph.getHtml(): OfficeExtension.ClientResult<string>",
        examples: [],
      },
    ],
  },
  {
    objName: "Word.Comment",
    apiList: [
      {
        name: "Word.Comment.authorEmail",
        description: "Get the email of the comment's author",
        kind: "Property",
        signature: "Word.Comment.authorEmail: string",
        examples: [],
      },
      {
        name: "Word.Comment.authorName",
        description: "Gets the name of the comment's author.",
        kind: "Property",
        signature: "Word.Comment.authorName: string",
        examples: [],
      },
      {
        name: "Word.Comment.content",
        description: "get or set the comment's content as plain text.",
        kind: "Property",
        signature: "Word.Comment.content",
        examples: [
          "const comment = context.document.getSelection().getComments().getFirst();\n comment.content = text;\n",
        ],
      },
      {
        name: "Word.Comment.creationDate",
        description: "Gets the creation date of the comment",
        kind: "Property",
        signature: "Word.Comment.creationDate: string",
        examples: [
          'const comment = context.document.getSelection().getComments().getFirst();\n comment.load("creationDate");\n',
        ],
      },
      {
        name: "Word.Comment.replies",
        description: "Gets the collection of reply objects associated with the comment.",
        kind: "Property",
        signature: "Word.Comment.replies: Word.CommentReplyCollection",
        examples: [],
      },
      {
        name: "Word.Comment.resolved",
        description:
          "Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.",
        kind: "Property",
        signature: "Word.Comment.resolved: boolean",
        examples: [
          "const comment = context.document.getSelection().getComments().getFirst();\n comment.resolved = true;\n",
        ],
      },
      {
        name: "Word.Comment.delete",
        description: "Deletes the comment and its replies.",
        kind: "Method",
        signature: "Word.Comment.delete: void",
        examples: [
          "const comment = context.document.getSelection().getComments().getFirst();\n comment.delete();\n",
        ],
      },
      {
        name: "Word.Comment.reply",
        description: "Reply the comment and its replies.",
        kind: "Method",
        signature: "Word.Comment.reply(replyText: string): Word.CommentReply",
        examples: [
          ' const comments = context.document.getSelection().getComments();\n  comments.load("items");\n  await context.sync();\n  const firstActiveComment = comments.items.find((item) => item.resolved !== true);\n   if (firstActiveComment) { \n   const reply = firstActiveComment.reply(text); \n    console.log("Reply added"); }',
        ],
      },
      {
        name: "Word.Comment.getRange",
        description: "Gets the range in the main document where the comment is on.",
        kind: "Method",
        signature: "Word.Comment.getRange(): Word.Range",
        examples: [
          ' const range = context.document.getSelection().getComments().getFirst().getRange(); \n range.load("text");\n  await context.sync();',
        ],
      },
    ],
  },
  {
    objName: "Excel.AllowEditRange",
    apiList: [
      {
        name: "Excel.AllowEditRange.address",
        description:
          "Specifies the range associated with the object. Worksheet protection must be disabled or paused for this method to work properly. If worksheet protection is enabled and not paused, this method throws an `AccessDenied` error and fails to set the range.",
        kind: "Property",
        signature: "Excel.AllowEditRange.address: string",
        examples: [],
      },
      {
        name: "Excel.AllowEditRange.isPasswordProtected",
        description: "Specifies if the object is password protected.",
        kind: "Property",
        signature: "Excel.AllowEditRange.isPasswordProtected: boolean",
        examples: [],
      },
      {
        name: "Excel.AllowEditRange.title",
        description:
          'Specifies the title of the object. Worksheet protection must be disabled or paused for this method to work properly. If worksheet protection is enabled and not paused, this method throws an `AccessDenied` error and fails to set the title. If there is already an existing `AllowEditRange` with the same string, or if the string is `null` or empty (""), then this method throws an `InvalidArgument` error and fails to set the title.',
        kind: "Property",
        signature: "Excel.AllowEditRange.title: string",
        examples: [],
      },
      {
        name: "Excel.AllowEditRange.delete",
        description:
          "Deletes the object from the `AllowEditRangeCollection`. Worksheet protection must be disabled or paused for this method to work properly. If worksheet protection is enabled and not paused, this method throws an `AccessDenied` error and fails the delete operation.",
        kind: "Method",
        signature: "Excel.AllowEditRange.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.AllowEditRange.pauseProtection",
        description:
          "Pauses worksheet protection for the object for the user in the current session. This method does nothing if worksheet protection isn't enabled or is already paused. If worksheet protection cannot be paused, this method throws an `UnsupportedOperation` error and fails to pause protection for the object. If the password is incorrect, then this method throws a `BadPassword` error and fails to pause protection for the object. If a password is supplied but the object does not require a password, the inputted password will be ignored and the operation will succeed.",
        kind: "Method",
        signature: "Excel.AllowEditRange.pauseProtection => (password?: string) => void",
        examples: [],
      },
      {
        name: "Excel.AllowEditRange.setPassword",
        description:
          'Changes the password associated with the object. Setting the password string as empty ("") or `null` will remove password protection from the object. Worksheet protection must be disabled or paused for this method to work properly. If worksheet protection is enabled and not paused, then this method throws an `AccessDenied` error and the set operation fails.',
        kind: "Method",
        signature: "Excel.AllowEditRange.setPassword => (password?: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.AllowEditRangeCollection",
    apiList: [
      {
        name: "Excel.AllowEditRangeCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.AllowEditRangeCollection.items: AllowEditRange[]",
        examples: [],
      },
      {
        name: "Excel.AllowEditRangeCollection.add",
        description:
          "Adds an `AllowEditRange` object to the worksheet. Worksheet protection must be disabled or paused for this method to work properly. If worksheet protection is enabled and not paused, then this method throws an `AccessDenied` error and the add operation fails.",
        kind: "Method",
        signature:
          "Excel.AllowEditRangeCollection.add => (title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions) => void",
        examples: [],
      },
      {
        name: "Excel.AllowEditRangeCollection.getCount",
        description: "Returns the number of `AllowEditRange` objects in the collection.",
        kind: "Method",
        signature:
          "Excel.AllowEditRangeCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.AllowEditRangeCollection.getItem",
        description: "Gets the `AllowEditRange` object by its title.",
        kind: "Method",
        signature:
          "Excel.AllowEditRangeCollection.getItem => (key: string) => Excel.AllowEditRange",
        examples: [],
      },
      {
        name: "Excel.AllowEditRangeCollection.getItemAt",
        description: "Returns an `AllowEditRange` object by its index in the collection.",
        kind: "Method",
        signature:
          "Excel.AllowEditRangeCollection.getItemAt => (index: number) => Excel.AllowEditRange",
        examples: [],
      },
      {
        name: "Excel.AllowEditRangeCollection.pauseProtection",
        description:
          "Pauses worksheet protection for all `AllowEditRange` objects found in this worksheet that have the given password for the user in the current session. This method does nothing if worksheet protection isn't enabled or is paused. If worksheet protection cannot be paused, this method throws an `UnsupportedOperation` error and fails to pause protection for the range. If the password does not match any `AllowEditRange` objects in the collection, then this method throws a `BadPassword` error and fails to pause protection for any range in the collection.",
        kind: "Method",
        signature: "Excel.AllowEditRangeCollection.pauseProtection => (password: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.AllowEditRangeOptions",
    apiList: [
      {
        name: "Excel.AllowEditRangeOptions.password",
        description: "The password associated with the `AllowEditRange`.",
        kind: "Property",
        signature: "Excel.AllowEditRangeOptions.password: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Application",
    apiList: [
      {
        name: "Excel.Application.calculationEngineVersion",
        description:
          "Returns the Excel calculation engine version used for the last full recalculation.",
        kind: "Property",
        signature: "Excel.Application.calculationEngineVersion: number",
        examples: [],
      },
      {
        name: "Excel.Application.calculationMode",
        description:
          "Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.",
        kind: "Property",
        signature:
          'Excel.Application.calculationMode: Excel.CalculationMode | "Automatic" | "AutomaticExceptTables" | "Manual"',
        examples: [
          '[rangeToGet.values, app.calculationMode, rangeToGet.values].join("\\n");',
          "application.calculationMode;",
          "workbook.application.calculationMode = Excel.CalculationMode.manual;",
          '"Current calculation mode: " + workbook.application.calculationMode;',
        ],
      },
      {
        name: "Excel.Application.calculationState",
        description:
          "Returns the calculation state of the application. See `Excel.CalculationState` for details.",
        kind: "Property",
        signature:
          'Excel.Application.calculationState: CalculationState | "Done" | "Calculating" | "Pending"',
        examples: [],
      },
      {
        name: "Excel.Application.cultureInfo",
        description:
          "Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.",
        kind: "Property",
        signature: "Excel.Application.cultureInfo: Excel.CultureInfo",
        examples: [
          "const systemDecimalSeparator = workbook.application.cultureInfo.numberFormat.numberDecimalSeparator;",
          "const systemThousandsSeparator = workbook.application.cultureInfo.numberFormat.numberGroupSeparator;",
          "const systemLongDatePattern = workbook.application.cultureInfo.datetimeFormat.longDatePattern;",
          "const systemShortDatePattern = workbook.application.cultureInfo.datetimeFormat.shortDatePattern;",
          "const systemDateSeparator = workbook.application.cultureInfo.datetimeFormat.dateSeparator;",
          "const systemLongTimePattern = workbook.application.cultureInfo.datetimeFormat.longTimePattern;",
          "const systemTimeSeparator = workbook.application.cultureInfo.datetimeFormat.timeSeparator;",
        ],
      },
      {
        name: "Excel.Application.decimalSeparator",
        description:
          "Gets the string used as the decimal separator for numeric values. This is based on the local Excel settings.",
        kind: "Property",
        signature: "Excel.Application.decimalSeparator: string",
        examples: ["const localDecimalSeparator = workbook.application.decimalSeparator;"],
      },
      {
        name: "Excel.Application.iterativeCalculation",
        description:
          "Returns the iterative calculation settings. In Excel on Windows and Mac, the settings will apply to the Excel Application. In Excel on the web and other platforms, the settings will apply to the active workbook.",
        kind: "Property",
        signature: "Excel.Application.iterativeCalculation: IterativeCalculation",
        examples: [],
      },
      {
        name: "Excel.Application.thousandsSeparator",
        description:
          "Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on the local Excel settings.",
        kind: "Property",
        signature: "Excel.Application.thousandsSeparator: string",
        examples: ["const localThousandsSeparator = workbook.application.thousandsSeparator;"],
      },
      {
        name: "Excel.Application.useSystemSeparators",
        description:
          "Specifies if the system separators of Excel are enabled. System separators include the decimal separator and thousands separator.",
        kind: "Property",
        signature: "Excel.Application.useSystemSeparators: boolean",
        examples: [],
      },
      {
        name: "Excel.Application.calculate",
        description: "Recalculate all currently opened workbooks in Excel.",
        kind: "Method",
        signature: "Excel.Application.calculate(calculationType: Excel.CalculationType): void",
        examples: [
          "workbook.application.calculate(Excel.CalculationType.full);",
          'workbook.application.calculate("Full");',
          "workbook.application.calculate(Excel.CalculationType.recalculate);",
        ],
      },
      {
        name: "Excel.Application.createWorkbook",
        description:
          "Creates a new hidden workbook by using an optional base64-encoded .xlsx file.",
        kind: "Method",
        signature:
          "Excel.Application.createWorkbook => (base64File?: string) => Excel.WorkbookCreated",
        examples: [],
      },
      {
        name: "Excel.Application.suspendApiCalculationUntilNextSync",
        description:
          "Suspends calculation until the next `context.sync()` is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.",
        kind: "Method",
        signature: "Excel.Application.suspendApiCalculationUntilNextSync => () => void",
        examples: ["app.suspendApiCalculationUntilNextSync();"],
      },
      {
        name: "Excel.Application.suspendScreenUpdatingUntilNextSync",
        description:
          "Suspends screen updating until the next `context.sync()` is called. **Note**: Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop). Repeated calls will cause the Excel window to flicker.",
        kind: "Method",
        signature: "Excel.Application.suspendScreenUpdatingUntilNextSync => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ArrayCellValue",
    apiList: [
      {
        name: "Excel.ArrayCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.ArrayCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.ArrayCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.ArrayCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.ArrayCellValue.elements",
        description:
          "Represents the elements of the array. May not directly contain an `ArrayCellValue`.",
        kind: "Property",
        signature: "Excel.ArrayCellValue.elements: CellValue[][]",
        examples: [],
      },
      {
        name: "Excel.ArrayCellValue.referencedValues",
        description:
          "Represents the cell values which are referenced within `ArrayCellValue.elements`.",
        kind: "Property",
        signature: "Excel.ArrayCellValue.referencedValues: ReferencedValue[]",
        examples: [],
      },
      {
        name: "Excel.ArrayCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature:
          'Excel.ArrayCellValue.type: CellValueType.array | ReferenceValueType.array | "Array"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.AutoFilter",
    apiList: [
      {
        name: "Excel.AutoFilter.criteria",
        description: "An array that holds all the filter criteria in the autofiltered range.",
        kind: "Property",
        signature: "Excel.AutoFilter.criteria: FilterCriteria[]",
        examples: [],
      },
      {
        name: "Excel.AutoFilter.enabled",
        description: "Specifies if the AutoFilter is enabled.",
        kind: "Property",
        signature: "Excel.AutoFilter.enabled: boolean",
        examples: [],
      },
      {
        name: "Excel.AutoFilter.isDataFiltered",
        description: "Specifies if the AutoFilter has filter criteria.",
        kind: "Property",
        signature: "Excel.AutoFilter.isDataFiltered: boolean",
        examples: [],
      },
      {
        name: "Excel.AutoFilter.apply",
        description:
          "Applies the AutoFilter to a range. This filters the column if column index and filter criteria are specified.",
        kind: "Method",
        signature:
          "Excel.AutoFilter.apply(range: string | Excel.Range, columnIndex?: number, criteria?: Excel.FilterCriteria) => void",
        examples: [
          'activeTable.autoFilter.apply(activeTable.getRange(), 2, {\n    filterOn: Excel.FilterOn.values,\n    values: ["Restaurant", "Groceries"],\n  });',
          "activeTable.autoFilter.apply(activeTable.getRange(), 3, {\n    filterOn: Excel.FilterOn.dynamic,\n    dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage,\n  });",
          'activeWorksheet.autoFilter.apply(farmData, 3, {\n    criterion1: "50",\n    filterOn: Excel.FilterOn.topPercent,\n  });',
          'activeWorksheet.autoFilter.apply(farmData, 1, {\n    criterion1: "=*e",\n    filterOn: Excel.FilterOn.custom,\n  });',
        ],
      },
      {
        name: "Excel.AutoFilter.clearColumnCriteria",
        description: "Clears the column filter criteria of the AutoFilter.",
        kind: "Method",
        signature: "Excel.AutoFilter.clearColumnCriteria(columnIndex: number) => void",
        examples: ["activeWorksheet.autoFilter.clearColumnCriteria(3);"],
      },
      {
        name: "Excel.AutoFilter.clearCriteria",
        description: "Clears the filter criteria and sort state of the AutoFilter.",
        kind: "Method",
        signature: "Excel.AutoFilter.clearCriteria => () => void",
        examples: [],
      },
      {
        name: "Excel.AutoFilter.getRange",
        description:
          "Returns the `Range` object that represents the range to which the AutoFilter applies.",
        kind: "Method",
        signature: "Excel.AutoFilter.getRange => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.AutoFilter.getRangeOrNullObject",
        description:
          "Returns the `Range` object that represents the range to which the AutoFilter applies. If there is no `Range` object associated with the AutoFilter, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.AutoFilter.getRangeOrNullObject => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.AutoFilter.reapply",
        description: "Applies the specified AutoFilter object currently on the range.",
        kind: "Method",
        signature: "Excel.AutoFilter.reapply() => void",
        examples: ["activeWorksheet.autoFilter.reapply();"],
      },
      {
        name: "Excel.AutoFilter.remove",
        description: "Removes the AutoFilter for the range.",
        kind: "Method",
        signature: "Excel.AutoFilter.remove() => void",
        examples: ["activeWorksheet.autoFilter.remove();"],
      },
    ],
  },
  {
    objName: "Excel.BasicDataValidation",
    apiList: [
      {
        name: "Excel.BasicDataValidation.formula1",
        description:
          'Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand. For example, setting formula1 to 10 and operator to GreaterThan means that valid data for the range must be greater than 10. When setting the value, it can be passed in as a number, a range object, or a string formula (where the string is either a stringified number, a cell reference like "=A1", or a formula like "=MIN(A1, B1)"). When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.',
        kind: "Property",
        signature: "Excel.BasicDataValidation.formula1: string | number | Range",
        examples: [],
      },
      {
        name: "Excel.BasicDataValidation.formula2",
        description:
          'With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan. When setting the value, it can be passed in as a number, a range object, or a string formula (where the string is either a stringified number, a cell reference like "=A1", or a formula like "=MIN(A1, B1)"). When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.',
        kind: "Property",
        signature: "Excel.BasicDataValidation.formula2: string | number | Range",
        examples: [],
      },
      {
        name: "Excel.BasicDataValidation.operator",
        description: "The operator to use for validating the data.",
        kind: "Property",
        signature:
          'Excel.BasicDataValidation.operator: "Between" | "GreaterThan" | "GreaterThanOrEqualTo" | "LessThan" | "LessThanOrEqualTo" | DataValidationOperator | "NotBetween" | "EqualTo" | "NotEqualTo"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Binding",
    apiList: [
      {
        name: "Excel.Binding.id",
        description: "Represents the binding identifier.",
        kind: "Property",
        signature: "Excel.Binding.id: string",
        examples: [],
      },
      {
        name: "Excel.Binding.type",
        description: "Returns the type of the binding. See `Excel.BindingType` for details.",
        kind: "Property",
        signature: 'Excel.Binding.type: Excel.BindingType | "Range" | "Table" | "Text"',
        examples: ["binding.type;"],
      },
      {
        name: "Excel.Binding.delete",
        description: "Deletes the binding.",
        kind: "Method",
        signature: "Excel.Binding.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.Binding.getRange",
        description:
          "Returns the range represented by the binding. Will throw an error if the binding is not of the correct type.",
        kind: "Method",
        signature: "Excel.Binding.getRange() => Excel.Range",
        examples: ["const range = binding.getRange();"],
      },
      {
        name: "Excel.Binding.getTable",
        description:
          "Returns the table represented by the binding. Will throw an error if the binding is not of the correct type.",
        kind: "Method",
        signature: "Excel.Binding.getTable() => Excel.Table",
        examples: ["const table = binding.getTable();"],
      },
      {
        name: "Excel.Binding.getText",
        description:
          "Returns the text represented by the binding. Will throw an error if the binding is not of the correct type.",
        kind: "Method",
        signature: "Excel.Binding.getText() => OfficeExtension.ClientResult<string>",
        examples: ["const text = binding.getText();"],
      },
    ],
  },
  {
    objName: "Excel.BindingCollection",
    apiList: [
      {
        name: "Excel.BindingCollection.count",
        description: "Returns the number of bindings in the collection.",
        kind: "Property",
        signature: "Excel.BindingCollection.count: number",
        examples: ["const lastPosition = workbook.bindings.count - 1;"],
      },
      {
        name: "Excel.BindingCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.BindingCollection.items: Binding[]",
        examples: [],
      },
      {
        name: "Excel.BindingCollection.add",
        description: "Add a new binding to a particular Range.",
        kind: "Method",
        signature:
          'Excel.BindingCollection.add => { (range: string | Range, bindingType: BindingType, id: string): Binding; (range: string | Range, bindingType: "Table" | "Text" | "Range", id: string): Binding; (range: Range | string, bindingType: string, id: string): Excel.Binding; }',
        examples: [],
      },
      {
        name: "Excel.BindingCollection.addFromNamedItem",
        description:
          "Add a new binding based on a named item in the workbook. If the named item references to multiple areas, the `InvalidReference` error will be returned.",
        kind: "Method",
        signature:
          'Excel.BindingCollection.addFromNamedItem => { (name: string, bindingType: BindingType, id: string): Binding; (name: string, bindingType: "Table" | "Text" | "Range", id: string): Binding; (name: string, bindingType: string, id: string): Excel.Binding; }',
        examples: [],
      },
      {
        name: "Excel.BindingCollection.addFromSelection",
        description:
          "Add a new binding based on the current selection. If the selection has multiple areas, the `InvalidReference` error will be returned.",
        kind: "Method",
        signature:
          'Excel.BindingCollection.addFromSelection => { (bindingType: BindingType, id: string): Binding; (bindingType: "Table" | "Text" | "Range", id: string): Binding; (bindingType: string, id: string): Excel.Binding; }',
        examples: [],
      },
      {
        name: "Excel.BindingCollection.getCount",
        description: "Gets the number of bindings in the collection.",
        kind: "Method",
        signature: "Excel.BindingCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.BindingCollection.getItem",
        description: "Gets a binding object by ID.",
        kind: "Method",
        signature: "Excel.BindingCollection.getItem => (id: string) => Excel.Binding",
        examples: [],
      },
      {
        name: "Excel.BindingCollection.getItemAt",
        description: "Gets a binding object based on its position in the items array.",
        kind: "Method",
        signature: "Excel.BindingCollection.getItemAt(index: number) => Excel.Binding",
        examples: [
          "const binding = workbook.bindings.getItemAt(0);",
          "const binding = workbook.bindings.getItemAt(lastPosition);",
        ],
      },
    ],
  },
  {
    objName: "Excel.BlockedErrorCellValue",
    apiList: [
      {
        name: "Excel.BlockedErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.BlockedErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.BlockedErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.BlockedErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.BlockedErrorCellValue.errorSubType",
        description: "Represents the type of `BlockedErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.BlockedErrorCellValue.errorSubType: BlockedErrorCellValueSubType | "Unknown" | "DataTypeRestrictedDomain" | "DataTypePrivacySetting" | "DataTypeUnsupportedApp" | "ExternalLinksGeneric" | "RichDataLinkDisabled" | "SignInError" | "NoLicense"',
        examples: [],
      },
      {
        name: "Excel.BlockedErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.BlockedErrorCellValue.errorType: ErrorCellValueType.blocked | "Blocked"',
        examples: [],
      },
      {
        name: "Excel.BlockedErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.BlockedErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.BooleanCellValue",
    apiList: [
      {
        name: "Excel.BooleanCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.BooleanCellValue.basicType: RangeValueType.boolean | "Boolean"',
        examples: [],
      },
      {
        name: "Excel.BooleanCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: "Excel.BooleanCellValue.basicValue: boolean",
        examples: [],
      },
      {
        name: "Excel.BooleanCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.BooleanCellValue.type: CellValueType.boolean | "Boolean"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.BusyErrorCellValue",
    apiList: [
      {
        name: "Excel.BusyErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.BusyErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.BusyErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.BusyErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.BusyErrorCellValue.errorSubType",
        description: "Represents the type of `BusyErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.BusyErrorCellValue.errorSubType: "Unknown" | "ExternalLinksGeneric" | BusyErrorCellValueSubType | "LoadingImage"',
        examples: [],
      },
      {
        name: "Excel.BusyErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.BusyErrorCellValue.errorType: ErrorCellValueType.busy | "Busy"',
        examples: [],
      },
      {
        name: "Excel.BusyErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.BusyErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CalcErrorCellValue",
    apiList: [
      {
        name: "Excel.CalcErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.CalcErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.CalcErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.CalcErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.CalcErrorCellValue.errorSubType",
        description: "Represents the type of `CalcErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.CalcErrorCellValue.errorSubType: "Unknown" | CalcErrorCellValueSubType | "ArrayOfArrays" | "ArrayOfRanges" | "EmptyArray" | "UnsupportedLifting" | "DataTableReferencedPendingFormula" | "TooManyCells" | "LambdaInCell" | "TooDeeplyNested" | "TextOverflow"',
        examples: [],
      },
      {
        name: "Excel.CalcErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.CalcErrorCellValue.errorType: ErrorCellValueType.calc | "Calc"',
        examples: [],
      },
      {
        name: "Excel.CalcErrorCellValue.functionName",
        description: "Represents the name of the function causing the error.",
        kind: "Property",
        signature: "Excel.CalcErrorCellValue.functionName: string",
        examples: [],
      },
      {
        name: "Excel.CalcErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.CalcErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CardLayoutPropertyReference",
    apiList: [
      {
        name: "Excel.CardLayoutPropertyReference.property",
        description: "Represents the name of the property referenced by the card layout.",
        kind: "Property",
        signature: "Excel.CardLayoutPropertyReference.property: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellPropertiesBorderLoadOptions",
    apiList: [
      {
        name: "Excel.CellPropertiesBorderLoadOptions.color",
        description: "Specifies whether to load the `color` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesBorderLoadOptions.color: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesBorderLoadOptions.style",
        description: "Specifies whether to load the `style` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesBorderLoadOptions.style: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesBorderLoadOptions.tintAndShade",
        description: "Specifies whether to load the `tintAndShade` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesBorderLoadOptions.tintAndShade: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesBorderLoadOptions.weight",
        description: "Specifies whether to load the `weight` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesBorderLoadOptions.weight: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellPropertiesFillLoadOptions",
    apiList: [
      {
        name: "Excel.CellPropertiesFillLoadOptions.color",
        description: "Specifies whether to load the `color` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFillLoadOptions.color: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFillLoadOptions.pattern",
        description: "Specifies whether to load the `pattern` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFillLoadOptions.pattern: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFillLoadOptions.patternColor",
        description: "Specifies whether to load the `patternColor` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFillLoadOptions.patternColor: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFillLoadOptions.patternTintAndShade",
        description: "Specifies whether to load the `patternTintAndShade` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFillLoadOptions.patternTintAndShade: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFillLoadOptions.tintAndShade",
        description: "Specifies whether to load the `tintAndShade` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFillLoadOptions.tintAndShade: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellPropertiesFontLoadOptions",
    apiList: [
      {
        name: "Excel.CellPropertiesFontLoadOptions.bold",
        description: "Specifies whether to load the `bold` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.bold: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.color",
        description: "Specifies whether to load the `color` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.color: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.italic",
        description: "Specifies whether to load the `italic` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.italic: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.name",
        description: "Specifies whether to load the `name` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.name: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.size",
        description: "Specifies whether to load the `size` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.size: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.strikethrough",
        description: "Specifies whether to load the `strikethrough` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.strikethrough: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.subscript",
        description: "Specifies whether to load the `subscript` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.subscript: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.superscript",
        description: "Specifies whether to load the `superscript` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.superscript: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.tintAndShade",
        description: "Specifies whether to load the `tintAndShade` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.tintAndShade: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFontLoadOptions.underline",
        description: "Specifies whether to load the `underline` property.",
        kind: "Property",
        signature: "Excel.CellPropertiesFontLoadOptions.underline: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellPropertiesFormatLoadOptions",
    apiList: [
      {
        name: "Excel.CellPropertiesFormatLoadOptions.autoIndent",
        description:
          "Specifies whether to load on the `autoIndent` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.autoIndent: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.borders",
        description: "Specifies whether to load on the `borders` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.borders: CellPropertiesBorderLoadOptions",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.fill",
        description: "Specifies whether to load on the `fill` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.fill: CellPropertiesFillLoadOptions",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.font",
        description: "Specifies whether to load on the `font` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.font: CellPropertiesFontLoadOptions",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.horizontalAlignment",
        description:
          "Specifies whether to load on the `horizontalAlignment` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.horizontalAlignment: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.indentLevel",
        description:
          "Specifies whether to load on the `indentLevel` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.indentLevel: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.protection",
        description:
          "Specifies whether to load on the `protection` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.protection: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.readingOrder",
        description:
          "Specifies whether to load on the `readingOrder` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.readingOrder: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.shrinkToFit",
        description:
          "Specifies whether to load on the `shrinkToFit` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.shrinkToFit: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.textOrientation",
        description:
          "Specifies whether to load on the `textOrientation` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.textOrientation: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.useStandardHeight",
        description:
          "Specifies whether to load on the `useStandardHeight` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.useStandardHeight: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.useStandardWidth",
        description:
          "Specifies whether to load on the `useStandardWidth` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.useStandardWidth: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.verticalAlignment",
        description:
          "Specifies whether to load on the `verticalAlignment` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.verticalAlignment: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesFormatLoadOptions.wrapText",
        description:
          "Specifies whether to load on the `wrapText` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesFormatLoadOptions.wrapText: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellPropertiesLoadOptions",
    apiList: [
      {
        name: "Excel.CellPropertiesLoadOptions.address",
        description: "Specifies whether to load on the `address` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesLoadOptions.address: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesLoadOptions.addressLocal",
        description:
          "Specifies whether to load on the `addressLocal` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesLoadOptions.addressLocal: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesLoadOptions.format",
        description: "Specifies whether to load on the `format` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesLoadOptions.format: CellPropertiesFormatLoadOptions",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesLoadOptions.hidden",
        description: "Specifies whether to load on the `hidden` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesLoadOptions.hidden: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesLoadOptions.hyperlink",
        description:
          "Specifies whether to load on the `hyperlink` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesLoadOptions.hyperlink: boolean",
        examples: [],
      },
      {
        name: "Excel.CellPropertiesLoadOptions.style",
        description: "Specifies whether to load on the `style` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.CellPropertiesLoadOptions.style: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellValueConditionalFormat",
    apiList: [
      {
        name: "Excel.CellValueConditionalFormat.format",
        description:
          "Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.",
        kind: "Property",
        signature: "Excel.CellValueConditionalFormat.format: Excel.ConditionalRangeFormat",
        examples: [
          'conditionalFormat.cellValue.format.font.color = "red";',
          'cellValueFormat.cellValue.format.font.color = "blue";',
          'cellValueFormat.cellValue.format.fill.color = "lightgreen";',
        ],
      },
      {
        name: "Excel.CellValueConditionalFormat.rule",
        description: "Specifies the rule object on this conditional format.",
        kind: "Property",
        signature: "Excel.CellValueConditionalFormat.rule: Excel.ConditionalCellValueRule",
        examples: [
          'conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };',
          'cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };',
        ],
      },
    ],
  },
  {
    objName: "Excel.CellValueExtraProperties",
    apiList: [
      {
        name: "Excel.CellValueExtraProperties.writable",
        description:
          "Represents whether this `CellValue` will be used to overwrite a cell. When false, APIs which would use this `CellValue` to overwrite a cell will instead ignore this value without throwing an error. The default value is true.",
        kind: "Property",
        signature: "Excel.CellValueExtraProperties.writable: boolean",
        examples: [],
      },
      {
        name: "Excel.CellValueExtraProperties.writableNote",
        description:
          "Represents an explanation about why `CellValue.writable` is specified as false. Note: This string is only available if `writable` is specified as false.",
        kind: "Property",
        signature: "Excel.CellValueExtraProperties.writableNote: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CellValueProviderAttributes",
    apiList: [
      {
        name: "Excel.CellValueProviderAttributes.description",
        description:
          "Represents the provider description property that is used in card view if no logo is specified. If a logo is specified, this will be used as tooltip text.",
        kind: "Property",
        signature: "Excel.CellValueProviderAttributes.description: string",
        examples: [],
      },
      {
        name: "Excel.CellValueProviderAttributes.logoSourceAddress",
        description:
          "Represents a URL used to download an image that will be used as a logo in card view.",
        kind: "Property",
        signature: "Excel.CellValueProviderAttributes.logoSourceAddress: string",
        examples: [],
      },
      {
        name: "Excel.CellValueProviderAttributes.logoTargetAddress",
        description:
          "Represents a URL that is the navigation target if the user clicks on the logo element in card view.",
        kind: "Property",
        signature: "Excel.CellValueProviderAttributes.logoTargetAddress: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChangedEventDetail",
    apiList: [
      {
        name: "Excel.ChangedEventDetail.valueAfter",
        description:
          "Represents the value after the change. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.",
        kind: "Property",
        signature: "Excel.ChangedEventDetail.valueAfter: any",
        examples: [],
      },
      {
        name: "Excel.ChangedEventDetail.valueAsJsonAfter",
        description:
          "Represents the type of value after the change. Unlike `valueAfter`, `valueAsJsonAfter` can represent all cell values, such as formatted number, web image, and entity data types.",
        kind: "Property",
        signature: "Excel.ChangedEventDetail.valueAsJsonAfter: CellValue",
        examples: [],
      },
      {
        name: "Excel.ChangedEventDetail.valueAsJsonBefore",
        description:
          "Represents the type of value before the change. Unlike `valueBefore`, `valueAsJsonBefore` can represent all cell values, such as formatted number, web image, and entity data types.",
        kind: "Property",
        signature: "Excel.ChangedEventDetail.valueAsJsonBefore: CellValue",
        examples: [],
      },
      {
        name: "Excel.ChangedEventDetail.valueBefore",
        description:
          "Represents the value before the change. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.",
        kind: "Property",
        signature: "Excel.ChangedEventDetail.valueBefore: any",
        examples: [],
      },
      {
        name: "Excel.ChangedEventDetail.valueTypeAfter",
        description: "Represents the type of value after the change.",
        kind: "Property",
        signature:
          'Excel.ChangedEventDetail.valueTypeAfter: RangeValueType | "Error" | "Unknown" | "Boolean" | "Double" | "Empty" | "String" | "Integer" | "RichValue"',
        examples: [],
      },
      {
        name: "Excel.ChangedEventDetail.valueTypeBefore",
        description: "Represents the type of value before the change.",
        kind: "Property",
        signature:
          'Excel.ChangedEventDetail.valueTypeBefore: RangeValueType | "Error" | "Unknown" | "Boolean" | "Double" | "Empty" | "String" | "Integer" | "RichValue"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChangeDirectionState",
    apiList: [
      {
        name: "Excel.ChangeDirectionState.deleteShiftDirection",
        description:
          "Represents the direction (such as up or to the left) that the remaining cells will shift when a cell or cells are deleted. Noteï¼š`insertShiftDirection` and `deleteShiftDirection` are exclusive and both enums can't have a value at the same time. If one has a value, then the other will return `undefined`.",
        kind: "Property",
        signature:
          'Excel.ChangeDirectionState.deleteShiftDirection: "Left" | "Up" | DeleteShiftDirection',
        examples: [],
      },
      {
        name: "Excel.ChangeDirectionState.insertShiftDirection",
        description:
          "Represents the direction (such as down or to the right) that the existing cells will shift when a new cell or cells are inserted. Noteï¼š`insertShiftDirection` and `deleteShiftDirection` are exclusive and both enums can't have a value at the same time. If one has a value, then the other will return `undefined`.",
        kind: "Property",
        signature:
          'Excel.ChangeDirectionState.insertShiftDirection: "Right" | "Down" | InsertShiftDirection',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Chart",
    apiList: [
      {
        name: "Excel.Chart.axes",
        description: "Represents chart axes.",
        kind: "Property",
        signature: "Excel.Chart.axes: Excel.ChartAxes",
        examples: [
          'activeChart.axes.categoryAxis.title.text = "Product";',
          'activeChart.axes.valueAxis.displayUnit = "Hundreds";',
          "activeChart.axes.valueAxis.majorGridlines.visible = false;",
          "activeChart.axes.valueAxis.maximum = 5;",
          "activeChart.axes.valueAxis.minimum = 0;",
          "activeChart.axes.valueAxis.majorUnit = 1;",
          "activeChart.axes.valueAxis.minorUnit = 0.2;",
          "let valueAxis = activeChart.axes.valueAxis;",
          "const axis = activeChart.axes.valueAxis;",
          "let axis = activeChart.axes.valueAxis;",
          'activeChart.axes.valueAxis.title.text = "Values";',
          'activeChart.axes.valueAxis.title.text = "Profits";',
          "activeChart.axes.valueAxis.title.textOrientation = 0;",
          "const gridlines = activeChart.axes.valueAxis.majorGridlines;",
          "activeChart.axes.valueAxis.majorGridlines.visible = true;",
        ],
      },
      {
        name: "Excel.Chart.categoryLabelLevel",
        description:
          "Specifies a chart category label level enumeration constant, referring to the level of the source category labels.",
        kind: "Property",
        signature: "Excel.Chart.categoryLabelLevel: number",
        examples: [],
      },
      {
        name: "Excel.Chart.chartType",
        description: "Specifies the type of the chart. See `Excel.ChartType` for details.",
        kind: "Property",
        signature:
          'Excel.Chart.chartType: Excel.ChartType | "Invalid" | "ColumnClustered" | "ColumnStacked" | "ColumnStacked100" | "3DColumnClustered" | "3DColumnStacked" | "3DColumnStacked100" | "BarClustered" | ... 73 more ... | "Funnel"',
        examples: ["activeChart.chartType = Excel.ChartType.barClustered;"],
      },
      {
        name: "Excel.Chart.dataLabels",
        description: "Represents the data labels on the chart.",
        kind: "Property",
        signature: "Excel.Chart.dataLabels: Excel.ChartDataLabels",
        examples: [
          "chart.dataLabels.format.font.size = 15;",
          'chart.dataLabels.format.font.color = "black";',
          "activeChart.dataLabels.showValue = true;",
          "activeChart.dataLabels.position = Excel.ChartDataLabelPosition.top;",
          "activeChart.dataLabels.showSeriesName = true;",
        ],
      },
      {
        name: "Excel.Chart.displayBlanksAs",
        description: "Specifies the way that blank cells are plotted on a chart.",
        kind: "Property",
        signature:
          'Excel.Chart.displayBlanksAs: ChartDisplayBlanksAs | "NotPlotted" | "Zero" | "Interplotted"',
        examples: [],
      },
      {
        name: "Excel.Chart.format",
        description: "Encapsulates the format properties for the chart area.",
        kind: "Property",
        signature: "Excel.Chart.format: ChartAreaFormat",
        examples: [],
      },
      {
        name: "Excel.Chart.height",
        description: "Specifies the height, in points, of the chart object.",
        kind: "Property",
        signature: "Excel.Chart.height: number",
        examples: ["activeChart.height = 200;", "chart.height = 300;"],
      },
      {
        name: "Excel.Chart.id",
        description: "The unique ID of chart.",
        kind: "Property",
        signature: "Excel.Chart.id: string",
        examples: [],
      },
      {
        name: "Excel.Chart.left",
        description:
          "The distance, in points, from the left side of the chart to the worksheet origin.",
        kind: "Property",
        signature: "Excel.Chart.left: number",
        examples: ["activeChart.left = 100;"],
      },
      {
        name: "Excel.Chart.legend",
        description: "Represents the legend for the chart.",
        kind: "Property",
        signature: "Excel.Chart.legend: Excel.ChartLegend",
        examples: [
          "chart.legend.position = Excel.ChartLegendPosition.right;",
          'chart.legend.format.fill.setSolidColor("white");',
          "activeChart.legend.visible = true;",
          'activeChart.legend.position = "Top";',
          "activeChart.legend.overlay = false;",
          "const legend = activeChart.legend;",
          "let font = activeChart.legend.format.font;",
          'chart.legend.position = "Right";',
        ],
      },
      {
        name: "Excel.Chart.name",
        description: "Specifies the name of a chart object.",
        kind: "Property",
        signature: "Excel.Chart.name: string",
        examples: [
          "activeChart.name;",
          'activeChart.name = "New Name";',
          "chart.name;",
          'bubbleChart.name = "Product Chart";',
        ],
      },
      {
        name: "Excel.Chart.pivotOptions",
        description: "Encapsulates the options for a pivot chart.",
        kind: "Property",
        signature: "Excel.Chart.pivotOptions: ChartPivotOptions",
        examples: [],
      },
      {
        name: "Excel.Chart.plotArea",
        description: "Represents the plot area for the chart.",
        kind: "Property",
        signature: "Excel.Chart.plotArea: ChartPlotArea",
        examples: [],
      },
      {
        name: "Excel.Chart.plotBy",
        description: "Specifies the way columns or rows are used as data series on the chart.",
        kind: "Property",
        signature: 'Excel.Chart.plotBy: "Columns" | "Rows" | ChartPlotBy',
        examples: [],
      },
      {
        name: "Excel.Chart.plotVisibleOnly",
        description:
          "True if only visible cells are plotted. False if both visible and hidden cells are plotted.",
        kind: "Property",
        signature: "Excel.Chart.plotVisibleOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.Chart.series",
        description: "Represents either a single series or collection of series in the chart.",
        kind: "Property",
        signature: "Excel.Chart.series: Excel.ChartSeriesCollection",
        examples: [
          'let newSeries = activeChart.series.add("2016");',
          "let seriesCollection = activeChart.series;",
          "let pointsCollection = activeChart.series.getItemAt(0).points;",
          "const points = activeChart.series.getItemAt(0).points;",
          "const pointsCollection = activeChart.series.getItemAt(0).points;",
          "const seriesCollection = activeChart.series;",
          "const firstSeries = activeChart.series.getItemAt(0);",
          'activeChart.series.getItemAt(0).name = "New Series Name";',
          "let series = chart.series;",
          "bubbleChart.series.getItemAt(0).delete();",
          "const newSeries = bubbleChart.series.add(dataRange.values[i][0], i);",
          'let newSeries = activeChart.series.add("Qtr2");',
        ],
      },
      {
        name: "Excel.Chart.seriesNameLevel",
        description:
          "Specifies a chart series name level enumeration constant, referring to the level of the source series names.",
        kind: "Property",
        signature: "Excel.Chart.seriesNameLevel: number",
        examples: [],
      },
      {
        name: "Excel.Chart.showAllFieldButtons",
        description: "Specifies whether to display all field buttons on a PivotChart.",
        kind: "Property",
        signature: "Excel.Chart.showAllFieldButtons: boolean",
        examples: [],
      },
      {
        name: "Excel.Chart.showDataLabelsOverMaximum",
        description:
          "Specifies whether to show the data labels when the value is greater than the maximum value on the value axis. If the value axis becomes smaller than the size of the data points, you can use this property to set whether to show the data labels. This property applies to 2-D charts only.",
        kind: "Property",
        signature: "Excel.Chart.showDataLabelsOverMaximum: boolean",
        examples: [],
      },
      {
        name: "Excel.Chart.style",
        description: "Specifies the chart style for the chart.",
        kind: "Property",
        signature: "Excel.Chart.style: number",
        examples: [],
      },
      {
        name: "Excel.Chart.title",
        description:
          "Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.",
        kind: "Property",
        signature: "Excel.Chart.title: Excel.ChartTitle",
        examples: [
          'chart.title.text = "Sales Data";',
          'activeChart.title.text = "Sales Data by Year";',
          "const title = activeChart.title;",
          'chart.title.text = "Bicycle Parts Quarterly Sales";',
          'activeChart.title.getSubstring(0, 7).font.color = "Yellow";',
          'activeChart.title.text = "My Chart";',
          "activeChart.title.visible = true;",
          "activeChart.title.overlay = true;",
        ],
      },
      {
        name: "Excel.Chart.top",
        description:
          "Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).",
        kind: "Property",
        signature: "Excel.Chart.top: number",
        examples: ["chart.top = 100;", "activeChart.top = 100;"],
      },
      {
        name: "Excel.Chart.width",
        description: "Specifies the width, in points, of the chart object.",
        kind: "Property",
        signature: "Excel.Chart.width: number",
        examples: ["activeChart.width = 200;", "chart.width = 500;"],
      },
      {
        name: "Excel.Chart.worksheet",
        description: "The worksheet containing the current chart.",
        kind: "Property",
        signature: "Excel.Chart.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.Chart.activate",
        description: "Activates the chart in the Excel UI.",
        kind: "Method",
        signature: "Excel.Chart.activate => () => void",
        examples: [],
      },
      {
        name: "Excel.Chart.delete",
        description: "Deletes the chart object.",
        kind: "Method",
        signature: "Excel.Chart.delete() => void",
        examples: ["activeChart.delete();"],
      },
      {
        name: "Excel.Chart.getDataRange",
        description: "Gets the data source of the whole chart.",
        kind: "Method",
        signature: "Excel.Chart.getDataRange => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.Chart.getDataRangeOrNullObject",
        description:
          "Gets the data source of the whole chart. If the data range is empty, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.Chart.getDataRangeOrNullObject => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.Chart.getDataTable",
        description:
          "Gets the data table on the chart. If the chart doesn't allow a data table, it will throw an exception.",
        kind: "Method",
        signature: "Excel.Chart.getDataTable => () => Excel.ChartDataTable",
        examples: [],
      },
      {
        name: "Excel.Chart.getDataTableOrNullObject",
        description:
          "Gets the data table on the chart. If the chart doesn't allow a data table, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Chart.getDataTableOrNullObject() => Excel.ChartDataTable",
        examples: ["const chartDataTable = activeChart.getDataTableOrNullObject();"],
      },
      {
        name: "Excel.Chart.getImage",
        description:
          "Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions. The aspect ratio is preserved as part of the resizing.",
        kind: "Method",
        signature:
          "Excel.Chart.getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>",
        examples: ["let imageAsString = activeChart.getImage();"],
      },
      {
        name: "Excel.Chart.setData",
        description: "Resets the source data for the chart.",
        kind: "Method",
        signature:
          "Excel.Chart.setData(sourceData: Excel.Range, seriesBy?: Excel.ChartSeriesBy): void",
        examples: ['activeChart.setData(sourceData, "Columns");'],
      },
      {
        name: "Excel.Chart.setPosition",
        description: "Positions the chart relative to cells on the worksheet.",
        kind: "Method",
        signature:
          "Excel.Chart.setPosition(startCell: string | Excel.Range, endCell?: string | Excel.Range) => void",
        examples: [
          'chart.setPosition("C2", null);',
          'chart.setPosition("A15", "E30");',
          'chart.setPosition("A22", "F35");',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartAreaFormat",
    apiList: [
      {
        name: "Excel.ChartAreaFormat.border",
        description:
          "Represents the border format of chart area, which includes color, linestyle, and weight.",
        kind: "Property",
        signature: "Excel.ChartAreaFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartAreaFormat.colorScheme",
        description: "Specifies the color scheme of the chart.",
        kind: "Property",
        signature:
          'Excel.ChartAreaFormat.colorScheme: ChartColorScheme | "ColorfulPalette1" | "ColorfulPalette2" | "ColorfulPalette3" | "ColorfulPalette4" | "MonochromaticPalette1" | "MonochromaticPalette2" | ... 10 more ... | "MonochromaticPalette13"',
        examples: [],
      },
      {
        name: "Excel.ChartAreaFormat.fill",
        description:
          "Represents the fill format of an object, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartAreaFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartAreaFormat.font",
        description:
          "Represents the font attributes (font name, font size, color, etc.) for the current object.",
        kind: "Property",
        signature: "Excel.ChartAreaFormat.font: ChartFont",
        examples: [],
      },
      {
        name: "Excel.ChartAreaFormat.roundedCorners",
        description: "Specifies if the chart area of the chart has rounded corners.",
        kind: "Property",
        signature: "Excel.ChartAreaFormat.roundedCorners: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartAxes",
    apiList: [
      {
        name: "Excel.ChartAxes.categoryAxis",
        description: "Represents the category axis in a chart.",
        kind: "Property",
        signature: "Excel.ChartAxes.categoryAxis: Excel.ChartAxis",
        examples: ['activeChart.axes.categoryAxis.title.text = "Product";'],
      },
      {
        name: "Excel.ChartAxes.seriesAxis",
        description: "Represents the series axis of a 3-D chart.",
        kind: "Property",
        signature: "Excel.ChartAxes.seriesAxis: ChartAxis",
        examples: [],
      },
      {
        name: "Excel.ChartAxes.valueAxis",
        description: "Represents the value axis in an axis.",
        kind: "Property",
        signature: "Excel.ChartAxes.valueAxis: Excel.ChartAxis",
        examples: [
          'activeChart.axes.valueAxis.displayUnit = "Hundreds";',
          "activeChart.axes.valueAxis.majorGridlines.visible = false;",
          "activeChart.axes.valueAxis.maximum = 5;",
          "activeChart.axes.valueAxis.minimum = 0;",
          "activeChart.axes.valueAxis.majorUnit = 1;",
          "activeChart.axes.valueAxis.minorUnit = 0.2;",
          "let valueAxis = activeChart.axes.valueAxis;",
          "const axis = activeChart.axes.valueAxis;",
          "let axis = activeChart.axes.valueAxis;",
          'activeChart.axes.valueAxis.title.text = "Values";',
          'activeChart.axes.valueAxis.title.text = "Profits";',
          "activeChart.axes.valueAxis.title.textOrientation = 0;",
          "const gridlines = activeChart.axes.valueAxis.majorGridlines;",
          "activeChart.axes.valueAxis.majorGridlines.visible = true;",
        ],
      },
      {
        name: "Excel.ChartAxes.getItem",
        description: "Returns the specific axis identified by type and group.",
        kind: "Method",
        signature:
          'Excel.ChartAxes.getItem => { (type: ChartAxisType, group?: ChartAxisGroup): ChartAxis; (type: "Value" | "Invalid" | "Category" | "Series", group?: "Primary" | "Secondary"): ChartAxis; (type: string, group?: string): Excel.ChartAxis; }',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartAxis",
    apiList: [
      {
        name: "Excel.ChartAxis.alignment",
        description:
          "Specifies the alignment for the specified axis tick label. See `Excel.ChartTextHorizontalAlignment` for detail.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.alignment: "Left" | "Center" | "Right" | ChartTickLabelAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.axisGroup",
        description:
          "Specifies the group for the specified axis. See `Excel.ChartAxisGroup` for details.",
        kind: "Property",
        signature: 'Excel.ChartAxis.axisGroup: ChartAxisGroup | "Primary" | "Secondary"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.baseTimeUnit",
        description: "Specifies the base unit for the specified category axis.",
        kind: "Property",
        signature: 'Excel.ChartAxis.baseTimeUnit: ChartAxisTimeUnit | "Days" | "Months" | "Years"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.categoryType",
        description: "Specifies the category axis type.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.categoryType: "Automatic" | ChartAxisCategoryType | "TextAxis" | "DateAxis"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.crosses",
        description:
          "[DEPRECATED; kept for back-compat with existing first-party solutions]. Please use `Position` instead. Specifies the specified axis where the other axis crosses. See `Excel.ChartAxisPosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.crosses: "Automatic" | "Custom" | ChartAxisPosition | "Maximum" | "Minimum"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.crossesAt",
        description:
          "[DEPRECATED; kept for back-compat with existing first-party solutions]. Please use `PositionAt` instead. Specifies the specified axis where the other axis crosses at. Set to this property should use `SetCrossesAt(double)` method.",
        kind: "Property",
        signature: "Excel.ChartAxis.crossesAt: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.customDisplayUnit",
        description:
          "Specifies the custom axis display unit value. To set this property, please use the `SetCustomDisplayUnit(double)` method.",
        kind: "Property",
        signature: "Excel.ChartAxis.customDisplayUnit: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.displayUnit",
        description:
          "Represents the axis display unit. See `Excel.ChartAxisDisplayUnit` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.displayUnit: Excel.ChartAxisDisplayUnit | "None" | "Hundreds" | "Thousands" | "TenThousands" | "HundredThousands" | "Millions" | "TenMillions" | "HundredMillions" | "Billions" | "Trillions" | "Custom"',
        examples: [
          'activeChart.axes.valueAxis.displayUnit = "Hundreds";',
          '"The vertical axis display unit is: " + valueAxis.displayUnit;',
        ],
      },
      {
        name: "Excel.ChartAxis.format",
        description:
          "Represents the formatting of a chart object, which includes line and font formatting.",
        kind: "Property",
        signature: "Excel.ChartAxis.format: ChartAxisFormat",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.height",
        description:
          "Specifies the height, in points, of the chart axis. Returns `null` if the axis is not visible.",
        kind: "Property",
        signature: "Excel.ChartAxis.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.isBetweenCategories",
        description: "Specifies if the value axis crosses the category axis between categories.",
        kind: "Property",
        signature: "Excel.ChartAxis.isBetweenCategories: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.left",
        description:
          "Specifies the distance, in points, from the left edge of the axis to the left of chart area. Returns `null` if the axis is not visible.",
        kind: "Property",
        signature: "Excel.ChartAxis.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.linkNumberFormat",
        description:
          "Specifies if the number format is linked to the cells. If `true`, the number format will change in the labels when it changes in the cells.",
        kind: "Property",
        signature: "Excel.ChartAxis.linkNumberFormat: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.logBase",
        description: "Specifies the base of the logarithm when using logarithmic scales.",
        kind: "Property",
        signature: "Excel.ChartAxis.logBase: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.majorGridlines",
        description:
          "Returns an object that represents the major gridlines for the specified axis.",
        kind: "Property",
        signature: "Excel.ChartAxis.majorGridlines: Excel.ChartGridlines",
        examples: [
          "activeChart.axes.valueAxis.majorGridlines.visible = false;",
          "const gridlines = activeChart.axes.valueAxis.majorGridlines;",
          "activeChart.axes.valueAxis.majorGridlines.visible = true;",
        ],
      },
      {
        name: "Excel.ChartAxis.majorTickMark",
        description:
          "Specifies the type of major tick mark for the specified axis. See `Excel.ChartAxisTickMark` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.majorTickMark: "None" | ChartAxisTickMark | "Cross" | "Inside" | "Outside"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.majorTimeUnitScale",
        description:
          "Specifies the major unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.majorTimeUnitScale: ChartAxisTimeUnit | "Days" | "Months" | "Years"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.majorUnit",
        description:
          "Represents the interval between two major tick marks. Can be set to a numeric value or an empty string. The returned value is always a number.",
        kind: "Property",
        signature: "Excel.ChartAxis.majorUnit: any",
        examples: ["activeChart.axes.valueAxis.majorUnit = 1;"],
      },
      {
        name: "Excel.ChartAxis.maximum",
        description:
          "Represents the maximum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.",
        kind: "Property",
        signature: "Excel.ChartAxis.maximum: any",
        examples: ["activeChart.axes.valueAxis.maximum = 5;", "axis.maximum;"],
      },
      {
        name: "Excel.ChartAxis.minimum",
        description:
          "Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.",
        kind: "Property",
        signature: "Excel.ChartAxis.minimum: any",
        examples: ["activeChart.axes.valueAxis.minimum = 0;"],
      },
      {
        name: "Excel.ChartAxis.minorGridlines",
        description:
          "Returns an object that represents the minor gridlines for the specified axis.",
        kind: "Property",
        signature: "Excel.ChartAxis.minorGridlines: ChartGridlines",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.minorTickMark",
        description:
          "Specifies the type of minor tick mark for the specified axis. See `Excel.ChartAxisTickMark` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.minorTickMark: "None" | ChartAxisTickMark | "Cross" | "Inside" | "Outside"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.minorTimeUnitScale",
        description:
          "Specifies the minor unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.minorTimeUnitScale: ChartAxisTimeUnit | "Days" | "Months" | "Years"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.minorUnit",
        description:
          "Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.",
        kind: "Property",
        signature: "Excel.ChartAxis.minorUnit: any",
        examples: ["activeChart.axes.valueAxis.minorUnit = 0.2;"],
      },
      {
        name: "Excel.ChartAxis.multiLevel",
        description: "Specifies if an axis is multilevel.",
        kind: "Property",
        signature: "Excel.ChartAxis.multiLevel: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.numberFormat",
        description: "Specifies the format code for the axis tick label.",
        kind: "Property",
        signature: "Excel.ChartAxis.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.offset",
        description:
          "Specifies the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.",
        kind: "Property",
        signature: "Excel.ChartAxis.offset: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.position",
        description:
          "Specifies the specified axis position where the other axis crosses. See `Excel.ChartAxisPosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.position: "Automatic" | "Custom" | ChartAxisPosition | "Maximum" | "Minimum"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.positionAt",
        description:
          "Specifies the axis position where the other axis crosses. You should use the `SetPositionAt(double)` method to set this property.",
        kind: "Property",
        signature: "Excel.ChartAxis.positionAt: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.reversePlotOrder",
        description: "Specifies if Excel plots data points from last to first.",
        kind: "Property",
        signature: "Excel.ChartAxis.reversePlotOrder: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.scaleType",
        description:
          "Specifies the value axis scale type. See `Excel.ChartAxisScaleType` for details.",
        kind: "Property",
        signature: 'Excel.ChartAxis.scaleType: ChartAxisScaleType | "Linear" | "Logarithmic"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.showDisplayUnitLabel",
        description: "Specifies if the axis display unit label is visible.",
        kind: "Property",
        signature: "Excel.ChartAxis.showDisplayUnitLabel: boolean",
        examples: ["axis.showDisplayUnitLabel = false;"],
      },
      {
        name: "Excel.ChartAxis.textOrientation",
        description:
          "Specifies the angle to which the text is oriented for the chart axis tick label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.",
        kind: "Property",
        signature: "Excel.ChartAxis.textOrientation: any",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.tickLabelPosition",
        description:
          "Specifies the position of tick-mark labels on the specified axis. See `Excel.ChartAxisTickLabelPosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.tickLabelPosition: "None" | ChartAxisTickLabelPosition | "NextToAxis" | "High" | "Low"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.tickLabelSpacing",
        description:
          "Specifies the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.",
        kind: "Property",
        signature: "Excel.ChartAxis.tickLabelSpacing: any",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.tickMarkSpacing",
        description: "Specifies the number of categories or series between tick marks.",
        kind: "Property",
        signature: "Excel.ChartAxis.tickMarkSpacing: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.title",
        description: "Represents the axis title.",
        kind: "Property",
        signature: "Excel.ChartAxis.title: Excel.ChartAxisTitle",
        examples: [
          'activeChart.axes.categoryAxis.title.text = "Product";',
          'activeChart.axes.valueAxis.title.text = "Values";',
          'activeChart.axes.valueAxis.title.text = "Profits";',
          "activeChart.axes.valueAxis.title.textOrientation = 0;",
        ],
      },
      {
        name: "Excel.ChartAxis.top",
        description:
          "Specifies the distance, in points, from the top edge of the axis to the top of chart area. Returns `null` if the axis is not visible.",
        kind: "Property",
        signature: "Excel.ChartAxis.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.type",
        description: "Specifies the axis type. See `Excel.ChartAxisType` for details.",
        kind: "Property",
        signature:
          'Excel.ChartAxis.type: "Value" | "Invalid" | ChartAxisType | "Category" | "Series"',
        examples: [],
      },
      {
        name: "Excel.ChartAxis.visible",
        description: "Specifies if the axis is visible.",
        kind: "Property",
        signature: "Excel.ChartAxis.visible: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.width",
        description:
          "Specifies the width, in points, of the chart axis. Returns `null` if the axis is not visible.",
        kind: "Property",
        signature: "Excel.ChartAxis.width: number",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.setCategoryNames",
        description: "Sets all the category names for the specified axis.",
        kind: "Method",
        signature: "Excel.ChartAxis.setCategoryNames => (sourceData: Range) => void",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.setCrossesAt",
        description:
          "[DEPRECATED; kept for back-compat with existing first-party solutions]. Please use `SetPositionAt` instead. Sets the specified axis where the other axis crosses at.",
        kind: "Method",
        signature: "Excel.ChartAxis.setCrossesAt => (value: number) => void",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.setCustomDisplayUnit",
        description: "Sets the axis display unit to a custom value.",
        kind: "Method",
        signature: "Excel.ChartAxis.setCustomDisplayUnit => (value: number) => void",
        examples: [],
      },
      {
        name: "Excel.ChartAxis.setPositionAt",
        description: "Sets the specified axis position where the other axis crosses.",
        kind: "Method",
        signature: "Excel.ChartAxis.setPositionAt => (value: number) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartAxisFormat",
    apiList: [
      {
        name: "Excel.ChartAxisFormat.fill",
        description: "Specifies chart fill formatting.",
        kind: "Property",
        signature: "Excel.ChartAxisFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartAxisFormat.font",
        description:
          "Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.",
        kind: "Property",
        signature: "Excel.ChartAxisFormat.font: ChartFont",
        examples: [],
      },
      {
        name: "Excel.ChartAxisFormat.line",
        description: "Specifies chart line formatting.",
        kind: "Property",
        signature: "Excel.ChartAxisFormat.line: ChartLineFormat",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartAxisTitle",
    apiList: [
      {
        name: "Excel.ChartAxisTitle.format",
        description: "Specifies the formatting of the chart axis title.",
        kind: "Property",
        signature: "Excel.ChartAxisTitle.format: ChartAxisTitleFormat",
        examples: [],
      },
      {
        name: "Excel.ChartAxisTitle.text",
        description: "Specifies the axis title.",
        kind: "Property",
        signature: "Excel.ChartAxisTitle.text: string",
        examples: [
          'activeChart.axes.categoryAxis.title.text = "Product";',
          'activeChart.axes.valueAxis.title.text = "Values";',
          'activeChart.axes.valueAxis.title.text = "Profits";',
        ],
      },
      {
        name: "Excel.ChartAxisTitle.textOrientation",
        description:
          "Specifies the angle to which the text is oriented for the chart axis title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.",
        kind: "Property",
        signature: "Excel.ChartAxisTitle.textOrientation: number",
        examples: ["activeChart.axes.valueAxis.title.textOrientation = 0;"],
      },
      {
        name: "Excel.ChartAxisTitle.visible",
        description: "Specifies if the axis title is visibile.",
        kind: "Property",
        signature: "Excel.ChartAxisTitle.visible: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartAxisTitle.setFormula",
        description:
          "A string value that represents the formula of chart axis title using A1-style notation.",
        kind: "Method",
        signature: "Excel.ChartAxisTitle.setFormula => (formula: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartAxisTitleFormat",
    apiList: [
      {
        name: "Excel.ChartAxisTitleFormat.border",
        description:
          "Specifies the chart axis title's border format, which includes color, linestyle, and weight.",
        kind: "Property",
        signature: "Excel.ChartAxisTitleFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartAxisTitleFormat.fill",
        description: "Specifies the chart axis title's fill formatting.",
        kind: "Property",
        signature: "Excel.ChartAxisTitleFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartAxisTitleFormat.font",
        description:
          "Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.",
        kind: "Property",
        signature: "Excel.ChartAxisTitleFormat.font: ChartFont",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartBinOptions",
    apiList: [
      {
        name: "Excel.ChartBinOptions.allowOverflow",
        description: "Specifies if bin overflow is enabled in a histogram chart or pareto chart.",
        kind: "Property",
        signature: "Excel.ChartBinOptions.allowOverflow: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartBinOptions.allowUnderflow",
        description: "Specifies if bin underflow is enabled in a histogram chart or pareto chart.",
        kind: "Property",
        signature: "Excel.ChartBinOptions.allowUnderflow: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartBinOptions.count",
        description: "Specifies the bin count of a histogram chart or pareto chart.",
        kind: "Property",
        signature: "Excel.ChartBinOptions.count: number",
        examples: [],
      },
      {
        name: "Excel.ChartBinOptions.overflowValue",
        description: "Specifies the bin overflow value of a histogram chart or pareto chart.",
        kind: "Property",
        signature: "Excel.ChartBinOptions.overflowValue: number",
        examples: [],
      },
      {
        name: "Excel.ChartBinOptions.type",
        description: "Specifies the bin's type for a histogram chart or pareto chart.",
        kind: "Property",
        signature:
          'Excel.ChartBinOptions.type: "Auto" | "Category" | ChartBinType | "BinWidth" | "BinCount"',
        examples: [],
      },
      {
        name: "Excel.ChartBinOptions.underflowValue",
        description: "Specifies the bin underflow value of a histogram chart or pareto chart.",
        kind: "Property",
        signature: "Excel.ChartBinOptions.underflowValue: number",
        examples: [],
      },
      {
        name: "Excel.ChartBinOptions.width",
        description: "Specifies the bin width value of a histogram chart or pareto chart.",
        kind: "Property",
        signature: "Excel.ChartBinOptions.width: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartBorder",
    apiList: [
      {
        name: "Excel.ChartBorder.color",
        description: "HTML color code representing the color of borders in the chart.",
        kind: "Property",
        signature: "Excel.ChartBorder.color: string",
        examples: ['chartDataTableFormat.border.color = "blue";'],
      },
      {
        name: "Excel.ChartBorder.lineStyle",
        description:
          "Represents the line style of the border. See `Excel.ChartLineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ChartBorder.lineStyle: "None" | "Automatic" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | ChartLineStyle | "Grey25" | "Grey50" | "Grey75" | "RoundDot"',
        examples: [],
      },
      {
        name: "Excel.ChartBorder.weight",
        description: "Represents weight of the border, in points.",
        kind: "Property",
        signature: "Excel.ChartBorder.weight: number",
        examples: [],
      },
      {
        name: "Excel.ChartBorder.clear",
        description: "Clear the border format of a chart element.",
        kind: "Method",
        signature: "Excel.ChartBorder.clear => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartBoxwhiskerOptions",
    apiList: [
      {
        name: "Excel.ChartBoxwhiskerOptions.quartileCalculation",
        description: "Specifies if the quartile calculation type of a box and whisker chart.",
        kind: "Property",
        signature:
          'Excel.ChartBoxwhiskerOptions.quartileCalculation: ChartBoxQuartileCalculation | "Inclusive" | "Exclusive"',
        examples: [],
      },
      {
        name: "Excel.ChartBoxwhiskerOptions.showInnerPoints",
        description: "Specifies if inner points are shown in a box and whisker chart.",
        kind: "Property",
        signature: "Excel.ChartBoxwhiskerOptions.showInnerPoints: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartBoxwhiskerOptions.showMeanLine",
        description: "Specifies if the mean line is shown in a box and whisker chart.",
        kind: "Property",
        signature: "Excel.ChartBoxwhiskerOptions.showMeanLine: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartBoxwhiskerOptions.showMeanMarker",
        description: "Specifies if the mean marker is shown in a box and whisker chart.",
        kind: "Property",
        signature: "Excel.ChartBoxwhiskerOptions.showMeanMarker: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartBoxwhiskerOptions.showOutlierPoints",
        description: "Specifies if outlier points are shown in a box and whisker chart.",
        kind: "Property",
        signature: "Excel.ChartBoxwhiskerOptions.showOutlierPoints: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartCollection",
    apiList: [
      {
        name: "Excel.ChartCollection.count",
        description: "Returns the number of charts in the worksheet.",
        kind: "Property",
        signature: "Excel.ChartCollection.count: number",
        examples: [
          '"charts: Count= " + charts.count;',
          'const lastPosition = workbook.worksheets.getItem("Sheet1").charts.count - 1;',
        ],
      },
      {
        name: "Excel.ChartCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ChartCollection.items: Chart[]",
        examples: [],
      },
      {
        name: "Excel.ChartCollection.add",
        description: "Inserts Creates or add a new chart.",
        kind: "Method",
        signature:
          "Excel.ChartCollection.add(type: Excel.ChartType, sourceData: Excel.Range, seriesBy?: Excel.ChartSeriesBy): Excel.Chart",
        examples: [
          "let chart = activeWorksheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.auto);",
          'let chart = activeWorksheet.charts.add(Excel.ChartType.columnStacked, activeWorksheet.getRange("B3:C5"));',
          'const chart = workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");',
          "activeWorksheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);",
          'let chart = activeWorksheet.charts.add("XYScatterSmooth", dataRange, "Auto");',
          "const bubbleChart = activeWorksheet.charts.add(Excel.ChartType.bubble, valueRange);",
          'let chart = sheet.charts.add("Line", dataRange, Excel.ChartSeriesBy.rows);',
          'let chart = activeWorksheet.charts.add(Excel.ChartType.line, dataRange, "Auto");',
        ],
      },
      {
        name: "Excel.ChartCollection.getCount",
        description: "Returns the number of charts in the worksheet.",
        kind: "Method",
        signature: "Excel.ChartCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.ChartCollection.getItem",
        description:
          "Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.",
        kind: "Method",
        signature: "Excel.ChartCollection.getItem(name: string) => Excel.Chart",
        examples: [
          'const activeChart = activeWorksheet.charts.getItem("Chart1");',
          'const activeChart = activeWorksheet.charts.getItem("SalesChart");',
          'const activeChart = activeWorksheet.charts.getItem("Sales Chart");',
          'const activeChart = activeWorksheet.charts.getItem("Product Chart");',
        ],
      },
      {
        name: "Excel.ChartCollection.getItemAt",
        description: "Gets a chart based on its position in the collection.",
        kind: "Method",
        signature: "Excel.ChartCollection.getItemAt(index: number) => Excel.Chart",
        examples: [
          'const chart = workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartDataLabel",
    apiList: [
      {
        name: "Excel.ChartDataLabel.autoText",
        description:
          "Specifies if the data label automatically generates appropriate text based on context.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.autoText: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.format",
        description: "Represents the format of chart data label.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.format: ChartDataLabelFormat",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.formula",
        description:
          "String value that represents the formula of chart data label using A1-style notation.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.formula: string",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.height",
        description:
          "Returns the height, in points, of the chart data label. Value is `null` if the chart data label is not visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.horizontalAlignment",
        description:
          "Represents the horizontal alignment for chart data label. See `Excel.ChartTextHorizontalAlignment` for details. This property is valid only when `TextOrientation` of data label is -90, 90, or 180.",
        kind: "Property",
        signature:
          'Excel.ChartDataLabel.horizontalAlignment: "Left" | "Center" | "Right" | "Justify" | "Distributed" | ChartTextHorizontalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.left",
        description:
          "Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Value is `null` if the chart data label is not visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.linkNumberFormat",
        description:
          "Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).",
        kind: "Property",
        signature: "Excel.ChartDataLabel.linkNumberFormat: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.numberFormat",
        description: "String value that represents the format code for data label.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.position",
        description:
          "Value that represents the position of the data label. See `Excel.ChartDataLabelPosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartDataLabel.position: "Left" | "Center" | "Right" | "Top" | "Bottom" | "None" | "Invalid" | ChartDataLabelPosition | "InsideEnd" | "InsideBase" | "OutsideEnd" | "BestFit" | "Callout"',
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.separator",
        description: "String representing the separator used for the data label on a chart.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.separator: string",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.showBubbleSize",
        description: "Specifies if the data label bubble size is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.showBubbleSize: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.showCategoryName",
        description: "Specifies if the data label category name is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.showCategoryName: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.showLegendKey",
        description: "Specifies if the data label legend key is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.showLegendKey: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.showPercentage",
        description: "Specifies if the data label percentage is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.showPercentage: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.showSeriesName",
        description: "Specifies if the data label series name is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.showSeriesName: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.showValue",
        description: "Specifies if the data label value is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.showValue: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.text",
        description: "String representing the text of the data label on a chart.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.text: string",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.textOrientation",
        description:
          "Represents the angle to which the text is oriented for the chart data label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.textOrientation: number",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.top",
        description:
          "Represents the distance, in points, from the top edge of chart data label to the top of chart area. Value is `null` if the chart data label is not visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.verticalAlignment",
        description:
          "Represents the vertical alignment of chart data label. See `Excel.ChartTextVerticalAlignment` for details. This property is valid only when `TextOrientation` of data label is 0.",
        kind: "Property",
        signature:
          'Excel.ChartDataLabel.verticalAlignment: "Center" | "Justify" | "Distributed" | "Top" | "Bottom" | ChartTextVerticalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartDataLabel.width",
        description:
          "Returns the width, in points, of the chart data label. Value is `null` if the chart data label is not visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabel.width: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartDataLabelFormat",
    apiList: [
      {
        name: "Excel.ChartDataLabelFormat.border",
        description: "Represents the border format, which includes color, linestyle, and weight.",
        kind: "Property",
        signature: "Excel.ChartDataLabelFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabelFormat.fill",
        description: "Represents the fill format of the current chart data label.",
        kind: "Property",
        signature: "Excel.ChartDataLabelFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabelFormat.font",
        description:
          "Represents the font attributes (such as font name, font size, and color) for a chart data label.",
        kind: "Property",
        signature: "Excel.ChartDataLabelFormat.font: Excel.ChartFont",
        examples: [
          "chart.dataLabels.format.font.size = 15;",
          'chart.dataLabels.format.font.color = "black";',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartDataLabels",
    apiList: [
      {
        name: "Excel.ChartDataLabels.autoText",
        description:
          "Specifies if data labels automatically generate appropriate text based on context.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.autoText: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.format",
        description:
          "Specifies the format of chart data labels, which includes fill and font formatting.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.format: Excel.ChartDataLabelFormat",
        examples: [
          "chart.dataLabels.format.font.size = 15;",
          'chart.dataLabels.format.font.color = "black";',
        ],
      },
      {
        name: "Excel.ChartDataLabels.horizontalAlignment",
        description:
          "Specifies the horizontal alignment for chart data label. See `Excel.ChartTextHorizontalAlignment` for details. This property is valid only when the `TextOrientation` of data label is 0.",
        kind: "Property",
        signature:
          'Excel.ChartDataLabels.horizontalAlignment: "Left" | "Center" | "Right" | "Justify" | "Distributed" | ChartTextHorizontalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.linkNumberFormat",
        description:
          "Specifies if the number format is linked to the cells. If `true`, the number format will change in the labels when it changes in the cells.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.linkNumberFormat: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.numberFormat",
        description: "Specifies the format code for data labels.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.position",
        description:
          "Value that represents the position of the data label. See `Excel.ChartDataLabelPosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartDataLabels.position: Excel.ChartDataLabelPosition | "Invalid" | "None" | "Center" | "InsideEnd" | "InsideBase" | "OutsideEnd" | "Left" | "Right" | "Top" | "Bottom" | "BestFit" | "Callout"',
        examples: ["activeChart.dataLabels.position = Excel.ChartDataLabelPosition.top;"],
      },
      {
        name: "Excel.ChartDataLabels.separator",
        description: "String representing the separator used for the data labels on a chart.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.separator: string",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.showBubbleSize",
        description: "Specifies if the data label bubble size is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.showBubbleSize: boolean",
        examples: ["newSeries.dataLabels.showBubbleSize = true;"],
      },
      {
        name: "Excel.ChartDataLabels.showCategoryName",
        description: "Specifies if the data label category name is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.showCategoryName: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.showLegendKey",
        description: "Specifies if the data label legend key is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.showLegendKey: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.showPercentage",
        description: "Specifies if the data label percentage is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.showPercentage: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.showSeriesName",
        description: "Specifies if the data label series name is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.showSeriesName: boolean",
        examples: [
          "activeChart.dataLabels.showSeriesName = true;",
          "newSeries.dataLabels.showSeriesName = true;",
        ],
      },
      {
        name: "Excel.ChartDataLabels.showValue",
        description: "Specifies if the data label value is visible.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.showValue: boolean",
        examples: [
          "activeChart.dataLabels.showValue = true;",
          "newSeries.dataLabels.showValue = false;",
        ],
      },
      {
        name: "Excel.ChartDataLabels.textOrientation",
        description:
          "Represents the angle to which the text is oriented for data labels. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.",
        kind: "Property",
        signature: "Excel.ChartDataLabels.textOrientation: number",
        examples: [],
      },
      {
        name: "Excel.ChartDataLabels.verticalAlignment",
        description:
          "Represents the vertical alignment of chart data label. See `Excel.ChartTextVerticalAlignment` for details. This property is valid only when `TextOrientation` of the data label is -90, 90, or 180.",
        kind: "Property",
        signature:
          'Excel.ChartDataLabels.verticalAlignment: "Center" | "Justify" | "Distributed" | "Top" | "Bottom" | ChartTextVerticalAlignment',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartDataTable",
    apiList: [
      {
        name: "Excel.ChartDataTable.format",
        description:
          "Represents the format of a chart data table, which includes fill, font, and border format.",
        kind: "Property",
        signature: "Excel.ChartDataTable.format: Excel.ChartDataTableFormat",
        examples: ["const chartDataTableFormat = chartDataTable.format;"],
      },
      {
        name: "Excel.ChartDataTable.showHorizontalBorder",
        description: "Specifies whether to display the horizontal border of the data table.",
        kind: "Property",
        signature: "Excel.ChartDataTable.showHorizontalBorder: boolean",
        examples: ["chartDataTable.showHorizontalBorder = false;"],
      },
      {
        name: "Excel.ChartDataTable.showLegendKey",
        description: "Specifies whether to show the legend key of the data table.",
        kind: "Property",
        signature: "Excel.ChartDataTable.showLegendKey: boolean",
        examples: ["chartDataTable.showLegendKey = true;"],
      },
      {
        name: "Excel.ChartDataTable.showOutlineBorder",
        description: "Specifies whether to display the outline border of the data table.",
        kind: "Property",
        signature: "Excel.ChartDataTable.showOutlineBorder: boolean",
        examples: ["chartDataTable.showOutlineBorder = true;"],
      },
      {
        name: "Excel.ChartDataTable.showVerticalBorder",
        description: "Specifies whether to display the vertical border of the data table.",
        kind: "Property",
        signature: "Excel.ChartDataTable.showVerticalBorder: boolean",
        examples: ["chartDataTable.showVerticalBorder = true;"],
      },
      {
        name: "Excel.ChartDataTable.visible",
        description: "Specifies whether to show the data table of the chart.",
        kind: "Property",
        signature: "Excel.ChartDataTable.visible: boolean",
        examples: ["chartDataTable.visible = true;"],
      },
    ],
  },
  {
    objName: "Excel.ChartDataTableFormat",
    apiList: [
      {
        name: "Excel.ChartDataTableFormat.border",
        description:
          "Represents the border format of chart data table, which includes color, line style, and weight.",
        kind: "Property",
        signature: "Excel.ChartDataTableFormat.border: Excel.ChartBorder",
        examples: ['chartDataTableFormat.border.color = "blue";'],
      },
      {
        name: "Excel.ChartDataTableFormat.fill",
        description:
          "Represents the fill format of an object, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartDataTableFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartDataTableFormat.font",
        description:
          "Represents the font attributes (such as font name, font size, and color) for the current object.",
        kind: "Property",
        signature: "Excel.ChartDataTableFormat.font: Excel.ChartFont",
        examples: [
          'chartDataTableFormat.font.color = "#B76E79";',
          'chartDataTableFormat.font.name = "Comic Sans";',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartErrorBars",
    apiList: [
      {
        name: "Excel.ChartErrorBars.endStyleCap",
        description: "Specifies if error bars have an end style cap.",
        kind: "Property",
        signature: "Excel.ChartErrorBars.endStyleCap: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartErrorBars.format",
        description: "Specifies the formatting type of the error bars.",
        kind: "Property",
        signature: "Excel.ChartErrorBars.format: ChartErrorBarsFormat",
        examples: [],
      },
      {
        name: "Excel.ChartErrorBars.include",
        description: "Specifies which parts of the error bars to include.",
        kind: "Property",
        signature:
          'Excel.ChartErrorBars.include: ChartErrorBarsInclude | "Both" | "MinusValues" | "PlusValues"',
        examples: [],
      },
      {
        name: "Excel.ChartErrorBars.type",
        description: "The type of range marked by the error bars.",
        kind: "Property",
        signature:
          'Excel.ChartErrorBars.type: "Percent" | "Custom" | ChartErrorBarsType | "FixedValue" | "StDev" | "StError"',
        examples: [],
      },
      {
        name: "Excel.ChartErrorBars.visible",
        description: "Specifies whether the error bars are displayed.",
        kind: "Property",
        signature: "Excel.ChartErrorBars.visible: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartErrorBarsFormat",
    apiList: [
      {
        name: "Excel.ChartErrorBarsFormat.line",
        description: "Represents the chart line formatting.",
        kind: "Property",
        signature: "Excel.ChartErrorBarsFormat.line: ChartLineFormat",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartFill",
    apiList: [
      {
        name: "Excel.ChartFill.clear",
        description: "Clears the fill color of a chart element.",
        kind: "Method",
        signature: "Excel.ChartFill.clear => () => void",
        examples: [],
      },
      {
        name: "Excel.ChartFill.getSolidColor",
        description: "Gets the uniform color fill formatting of a chart element.",
        kind: "Method",
        signature: "Excel.ChartFill.getSolidColor => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.ChartFill.setSolidColor",
        description: "Sets the fill formatting of a chart element to a uniform color.",
        kind: "Method",
        signature: "Excel.ChartFill.setSolidColor(color: string) => void",
        examples: [
          'chart.legend.format.fill.setSolidColor("white");',
          'point.format.fill.setSolidColor("red");',
          'points.getItemAt(0).format.fill.setSolidColor("8FBC8F");',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartFont",
    apiList: [
      {
        name: "Excel.ChartFont.bold",
        description: "Represents the bold status of font.",
        kind: "Property",
        signature: "Excel.ChartFont.bold: boolean",
        examples: ["title.format.font.bold = true;", "font.bold = true;"],
      },
      {
        name: "Excel.ChartFont.color",
        description:
          "HTML color code representation of the text color (e.g., #FF0000 represents Red).",
        kind: "Property",
        signature: "Excel.ChartFont.color: string",
        examples: [
          'chart.dataLabels.format.font.color = "black";',
          'chartDataTableFormat.font.color = "#B76E79";',
          'title.format.font.color = "#FF0000";',
          'font.color = "red";',
          'activeChart.title.getSubstring(0, 7).font.color = "Yellow";',
        ],
      },
      {
        name: "Excel.ChartFont.italic",
        description: "Represents the italic status of the font.",
        kind: "Property",
        signature: "Excel.ChartFont.italic: boolean",
        examples: ["title.format.font.italic = false;", "font.italic = true;"],
      },
      {
        name: "Excel.ChartFont.name",
        description: 'Font name (e.g., "Calibri")',
        kind: "Property",
        signature: "Excel.ChartFont.name: string",
        examples: [
          'chartDataTableFormat.font.name = "Comic Sans";',
          'title.format.font.name = "Calibri";',
          'font.name = "Calibri";',
        ],
      },
      {
        name: "Excel.ChartFont.size",
        description: "Size of the font (e.g., 11)",
        kind: "Property",
        signature: "Excel.ChartFont.size: number",
        examples: [
          "chart.dataLabels.format.font.size = 15;",
          "title.format.font.size = 12;",
          "font.size = 15;",
        ],
      },
      {
        name: "Excel.ChartFont.underline",
        description:
          "Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.",
        kind: "Property",
        signature: 'Excel.ChartFont.underline: Excel.ChartUnderlineStyle | "None" | "Single"',
        examples: ['title.format.font.underline = "None";', 'font.underline = "Single";'],
      },
    ],
  },
  {
    objName: "Excel.ChartFormatString",
    apiList: [
      {
        name: "Excel.ChartFormatString.font",
        description:
          "Represents the font attributes, such as font name, font size, and color of a chart characters object.",
        kind: "Property",
        signature: "Excel.ChartFormatString.font: Excel.ChartFont",
        examples: ['activeChart.title.getSubstring(0, 7).font.color = "Yellow";'],
      },
    ],
  },
  {
    objName: "Excel.ChartGridlines",
    apiList: [
      {
        name: "Excel.ChartGridlines.format",
        description: "Represents the formatting of chart gridlines.",
        kind: "Property",
        signature: "Excel.ChartGridlines.format: Excel.ChartGridlinesFormat",
        examples: ["gridlines.format.line.clear();", 'gridlines.format.line.color = "#FF0000";'],
      },
      {
        name: "Excel.ChartGridlines.visible",
        description: "Specifies if the axis gridlines are visible.",
        kind: "Property",
        signature: "Excel.ChartGridlines.visible: boolean",
        examples: [
          "activeChart.axes.valueAxis.majorGridlines.visible = false;",
          "activeChart.axes.valueAxis.majorGridlines.visible = true;",
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartGridlinesFormat",
    apiList: [
      {
        name: "Excel.ChartGridlinesFormat.line",
        description: "Represents chart line formatting.",
        kind: "Property",
        signature: "Excel.ChartGridlinesFormat.line: Excel.ChartLineFormat",
        examples: ["gridlines.format.line.clear();", 'gridlines.format.line.color = "#FF0000";'],
      },
    ],
  },
  {
    objName: "Excel.ChartLegend",
    apiList: [
      {
        name: "Excel.ChartLegend.format",
        description:
          "Represents the formatting of a chart legend, which includes fill and font formatting.",
        kind: "Property",
        signature: "Excel.ChartLegend.format: Excel.ChartLegendFormat",
        examples: [
          'chart.legend.format.fill.setSolidColor("white");',
          "let font = activeChart.legend.format.font;",
        ],
      },
      {
        name: "Excel.ChartLegend.height",
        description:
          "Specifies the height, in points, of the legend on the chart. Value is `null` if the legend is not visible.",
        kind: "Property",
        signature: "Excel.ChartLegend.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegend.left",
        description:
          "Specifies the left value, in points, of the legend on the chart. Value is `null` if the legend is not visible.",
        kind: "Property",
        signature: "Excel.ChartLegend.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegend.legendEntries",
        description: "Represents a collection of legendEntries in the legend.",
        kind: "Property",
        signature: "Excel.ChartLegend.legendEntries: ChartLegendEntryCollection",
        examples: [],
      },
      {
        name: "Excel.ChartLegend.overlay",
        description:
          "Specifies if the chart legend should overlap with the main body of the chart.",
        kind: "Property",
        signature: "Excel.ChartLegend.overlay: boolean",
        examples: ["activeChart.legend.overlay = false;"],
      },
      {
        name: "Excel.ChartLegend.position",
        description:
          "Specifies the position of the legend on the chart. See `Excel.ChartLegendPosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartLegend.position: Excel.ChartLegendPosition | "Invalid" | "Top" | "Bottom" | "Left" | "Right" | "Corner" | "Custom"',
        examples: [
          "chart.legend.position = Excel.ChartLegendPosition.right;",
          'activeChart.legend.position = "Top";',
          "legend.position;",
          'chart.legend.position = "Right";',
        ],
      },
      {
        name: "Excel.ChartLegend.showShadow",
        description: "Specifies if the legend has a shadow on the chart.",
        kind: "Property",
        signature: "Excel.ChartLegend.showShadow: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartLegend.top",
        description: "Specifies the top of a chart legend.",
        kind: "Property",
        signature: "Excel.ChartLegend.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegend.visible",
        description: "Specifies if the chart legend is visible.",
        kind: "Property",
        signature: "Excel.ChartLegend.visible: boolean",
        examples: ["activeChart.legend.visible = true;"],
      },
      {
        name: "Excel.ChartLegend.width",
        description:
          "Specifies the width, in points, of the legend on the chart. Value is `null` if the legend is not visible.",
        kind: "Property",
        signature: "Excel.ChartLegend.width: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartLegendEntry",
    apiList: [
      {
        name: "Excel.ChartLegendEntry.height",
        description: "Specifies the height of the legend entry on the chart legend.",
        kind: "Property",
        signature: "Excel.ChartLegendEntry.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntry.index",
        description: "Specifies the index of the legend entry in the chart legend.",
        kind: "Property",
        signature: "Excel.ChartLegendEntry.index: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntry.left",
        description: "Specifies the left value of a chart legend entry.",
        kind: "Property",
        signature: "Excel.ChartLegendEntry.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntry.top",
        description: "Specifies the top of a chart legend entry.",
        kind: "Property",
        signature: "Excel.ChartLegendEntry.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntry.visible",
        description: "Represents the visibility of a chart legend entry.",
        kind: "Property",
        signature: "Excel.ChartLegendEntry.visible: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntry.width",
        description: "Represents the width of the legend entry on the chart Legend.",
        kind: "Property",
        signature: "Excel.ChartLegendEntry.width: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartLegendEntryCollection",
    apiList: [
      {
        name: "Excel.ChartLegendEntryCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ChartLegendEntryCollection.items: ChartLegendEntry[]",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntryCollection.getCount",
        description: "Returns the number of legend entries in the collection.",
        kind: "Method",
        signature:
          "Excel.ChartLegendEntryCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.ChartLegendEntryCollection.getItemAt",
        description: "Returns a legend entry at the given index.",
        kind: "Method",
        signature:
          "Excel.ChartLegendEntryCollection.getItemAt => (index: number) => Excel.ChartLegendEntry",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartLegendFormat",
    apiList: [
      {
        name: "Excel.ChartLegendFormat.border",
        description: "Represents the border format, which includes color, linestyle, and weight.",
        kind: "Property",
        signature: "Excel.ChartLegendFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartLegendFormat.fill",
        description:
          "Represents the fill format of an object, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartLegendFormat.fill: Excel.ChartFill",
        examples: ['chart.legend.format.fill.setSolidColor("white");'],
      },
      {
        name: "Excel.ChartLegendFormat.font",
        description:
          "Represents the font attributes such as font name, font size, and color of a chart legend.",
        kind: "Property",
        signature: "Excel.ChartLegendFormat.font: Excel.ChartFont",
        examples: ["let font = activeChart.legend.format.font;"],
      },
    ],
  },
  {
    objName: "Excel.ChartLineFormat",
    apiList: [
      {
        name: "Excel.ChartLineFormat.color",
        description: "HTML color code representing the color of lines in the chart.",
        kind: "Property",
        signature: "Excel.ChartLineFormat.color: string",
        examples: [
          'gridlines.format.line.color = "#FF0000";',
          'line.color = "#FF0000";',
          '"The trendline color has been set to:" + line.color;',
        ],
      },
      {
        name: "Excel.ChartLineFormat.lineStyle",
        description: "Represents the line style. See `Excel.ChartLineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ChartLineFormat.lineStyle: "None" | "Automatic" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | ChartLineStyle | "Grey25" | "Grey50" | "Grey75" | "RoundDot"',
        examples: [],
      },
      {
        name: "Excel.ChartLineFormat.weight",
        description: "Represents weight of the line, in points.",
        kind: "Property",
        signature: "Excel.ChartLineFormat.weight: number",
        examples: [],
      },
      {
        name: "Excel.ChartLineFormat.clear",
        description: "Clears the line format of a chart element.",
        kind: "Method",
        signature: "Excel.ChartLineFormat.clear() => void",
        examples: ["gridlines.format.line.clear();"],
      },
    ],
  },
  {
    objName: "Excel.ChartMapOptions",
    apiList: [
      {
        name: "Excel.ChartMapOptions.labelStrategy",
        description: "Specifies the series map labels strategy of a region map chart.",
        kind: "Property",
        signature:
          'Excel.ChartMapOptions.labelStrategy: "None" | "BestFit" | ChartMapLabelStrategy | "ShowAll"',
        examples: [],
      },
      {
        name: "Excel.ChartMapOptions.level",
        description: "Specifies the series mapping level of a region map chart.",
        kind: "Property",
        signature:
          'Excel.ChartMapOptions.level: "Automatic" | ChartMapAreaLevel | "DataOnly" | "City" | "County" | "State" | "Country" | "Continent" | "World"',
        examples: [],
      },
      {
        name: "Excel.ChartMapOptions.projectionType",
        description: "Specifies the series projection type of a region map chart.",
        kind: "Property",
        signature:
          'Excel.ChartMapOptions.projectionType: "Automatic" | ChartMapProjectionType | "Mercator" | "Miller" | "Robinson" | "Albers"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartPivotOptions",
    apiList: [
      {
        name: "Excel.ChartPivotOptions.showAxisFieldButtons",
        description:
          'Specifies whether to display the axis field buttons on a PivotChart. The `showAxisFieldButtons` property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.',
        kind: "Property",
        signature: "Excel.ChartPivotOptions.showAxisFieldButtons: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartPivotOptions.showLegendFieldButtons",
        description: "Specifies whether to display the legend field buttons on a PivotChart.",
        kind: "Property",
        signature: "Excel.ChartPivotOptions.showLegendFieldButtons: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartPivotOptions.showReportFilterFieldButtons",
        description:
          "Specifies whether to display the report filter field buttons on a PivotChart.",
        kind: "Property",
        signature: "Excel.ChartPivotOptions.showReportFilterFieldButtons: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartPivotOptions.showValueFieldButtons",
        description: "Specifies whether to display the show value field buttons on a PivotChart.",
        kind: "Property",
        signature: "Excel.ChartPivotOptions.showValueFieldButtons: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartPlotArea",
    apiList: [
      {
        name: "Excel.ChartPlotArea.format",
        description: "Specifies the formatting of a chart plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.format: ChartPlotAreaFormat",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.height",
        description: "Specifies the height value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.insideHeight",
        description: "Specifies the inside height value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.insideHeight: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.insideLeft",
        description: "Specifies the inside left value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.insideLeft: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.insideTop",
        description: "Specifies the inside top value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.insideTop: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.insideWidth",
        description: "Specifies the inside width value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.insideWidth: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.left",
        description: "Specifies the left value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.position",
        description: "Specifies the position of a plot area.",
        kind: "Property",
        signature: 'Excel.ChartPlotArea.position: "Automatic" | "Custom" | ChartPlotAreaPosition',
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.top",
        description: "Specifies the top value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartPlotArea.width",
        description: "Specifies the width value of a plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotArea.width: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartPlotAreaFormat",
    apiList: [
      {
        name: "Excel.ChartPlotAreaFormat.border",
        description: "Specifies the border attributes of a chart plot area.",
        kind: "Property",
        signature: "Excel.ChartPlotAreaFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartPlotAreaFormat.fill",
        description:
          "Specifies the fill format of an object, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartPlotAreaFormat.fill: ChartFill",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartPoint",
    apiList: [
      {
        name: "Excel.ChartPoint.dataLabel",
        description: "Returns the data label of a chart point.",
        kind: "Property",
        signature: "Excel.ChartPoint.dataLabel: ChartDataLabel",
        examples: [],
      },
      {
        name: "Excel.ChartPoint.format",
        description: "Encapsulates the format properties chart point.",
        kind: "Property",
        signature: "Excel.ChartPoint.format: Excel.ChartPointFormat",
        examples: [
          'point.format.fill.setSolidColor("red");',
          'points.getItemAt(0).format.fill.setSolidColor("8FBC8F");',
        ],
      },
      {
        name: "Excel.ChartPoint.hasDataLabel",
        description:
          "Represents whether a data point has a data label. Not applicable for surface charts.",
        kind: "Property",
        signature: "Excel.ChartPoint.hasDataLabel: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartPoint.markerBackgroundColor",
        description:
          "HTML color code representation of the marker background color of a data point (e.g., #FF0000 represents Red).",
        kind: "Property",
        signature: "Excel.ChartPoint.markerBackgroundColor: string",
        examples: [],
      },
      {
        name: "Excel.ChartPoint.markerForegroundColor",
        description:
          "HTML color code representation of the marker foreground color of a data point (e.g., #FF0000 represents Red).",
        kind: "Property",
        signature: "Excel.ChartPoint.markerForegroundColor: string",
        examples: [],
      },
      {
        name: "Excel.ChartPoint.markerSize",
        description: "Represents marker size of a data point.",
        kind: "Property",
        signature: "Excel.ChartPoint.markerSize: number",
        examples: [],
      },
      {
        name: "Excel.ChartPoint.markerStyle",
        description:
          "Represents marker style of a chart data point. See `Excel.ChartMarkerStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ChartPoint.markerStyle: "None" | "Automatic" | "Dash" | "Dot" | "Invalid" | ChartMarkerStyle | "Square" | "Diamond" | "Triangle" | "X" | "Star" | "Circle" | "Plus" | "Picture"',
        examples: [],
      },
      {
        name: "Excel.ChartPoint.value",
        description: "Returns the value of a chart point.",
        kind: "Property",
        signature: "Excel.ChartPoint.value: any",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartPointFormat",
    apiList: [
      {
        name: "Excel.ChartPointFormat.border",
        description:
          "Represents the border format of a chart data point, which includes color, style, and weight information.",
        kind: "Property",
        signature: "Excel.ChartPointFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartPointFormat.fill",
        description:
          "Represents the fill format of a chart, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartPointFormat.fill: Excel.ChartFill",
        examples: [
          'point.format.fill.setSolidColor("red");',
          'points.getItemAt(0).format.fill.setSolidColor("8FBC8F");',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartPointsCollection",
    apiList: [
      {
        name: "Excel.ChartPointsCollection.count",
        description: "Returns the number of chart points in the series.",
        kind: "Property",
        signature: "Excel.ChartPointsCollection.count: number",
        examples: ['"points: Count= " + pointsCollection.count;'],
      },
      {
        name: "Excel.ChartPointsCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ChartPointsCollection.items: ChartPoint[]",
        examples: [],
      },
      {
        name: "Excel.ChartPointsCollection.getCount",
        description: "Returns the number of chart points in the series.",
        kind: "Method",
        signature:
          "Excel.ChartPointsCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.ChartPointsCollection.getItemAt",
        description: "Retrieve a point based on its position within the series.",
        kind: "Method",
        signature: "Excel.ChartPointsCollection.getItemAt(index: number) => Excel.ChartPoint",
        examples: [
          "let point = pointsCollection.getItemAt(2);",
          'points.getItemAt(0).format.fill.setSolidColor("8FBC8F");',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartSeries",
    apiList: [
      {
        name: "Excel.ChartSeries.axisGroup",
        description: "Specifies the group for the specified series.",
        kind: "Property",
        signature: 'Excel.ChartSeries.axisGroup: ChartAxisGroup | "Primary" | "Secondary"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.binOptions",
        description: "Encapsulates the bin options for histogram charts and pareto charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.binOptions: ChartBinOptions",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.boxwhiskerOptions",
        description: "Encapsulates the options for the box and whisker charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.boxwhiskerOptions: ChartBoxwhiskerOptions",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.bubbleScale",
        description:
          "This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.bubbleScale: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.chartType",
        description: "Represents the chart type of a series. See `Excel.ChartType` for details.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.chartType: ChartType | "Invalid" | "ColumnClustered" | "ColumnStacked" | "ColumnStacked100" | "3DColumnClustered" | "3DColumnStacked" | "3DColumnStacked100" | "BarClustered" | ... 73 more ... | "Funnel"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.dataLabels",
        description: "Represents a collection of all data labels in the series.",
        kind: "Property",
        signature: "Excel.ChartSeries.dataLabels: Excel.ChartDataLabels",
        examples: [
          "newSeries.dataLabels.showSeriesName = true;",
          "newSeries.dataLabels.showBubbleSize = true;",
          "newSeries.dataLabels.showValue = false;",
        ],
      },
      {
        name: "Excel.ChartSeries.doughnutHoleSize",
        description:
          "Represents the doughnut hole size of a chart series. Only valid on doughnut and doughnut exploded charts. Throws an `InvalidArgument` error on invalid charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.doughnutHoleSize: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.explosion",
        description:
          "Specifies the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).",
        kind: "Property",
        signature: "Excel.ChartSeries.explosion: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.filtered",
        description: "Specifies if the series is filtered. Not applicable for surface charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.filtered: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.firstSliceAngle",
        description:
          "Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.",
        kind: "Property",
        signature: "Excel.ChartSeries.firstSliceAngle: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.format",
        description:
          "Represents the formatting of a chart series, which includes fill and line formatting.",
        kind: "Property",
        signature: "Excel.ChartSeries.format: ChartSeriesFormat",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gapWidth",
        description:
          "Represents the gap width of a chart series. Only valid on bar and column charts, as well as specific classes of line and pie charts. Throws an invalid argument exception on invalid charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.gapWidth: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMaximumColor",
        description: "Specifies the color for maximum value of a region map chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.gradientMaximumColor: string",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMaximumType",
        description: "Specifies the type for maximum value of a region map chart series.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.gradientMaximumType: "Percent" | ChartGradientStyleType | "ExtremeValue" | "Number"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMaximumValue",
        description: "Specifies the maximum value of a region map chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.gradientMaximumValue: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMidpointColor",
        description: "Specifies the color for the midpoint value of a region map chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.gradientMidpointColor: string",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMidpointType",
        description: "Specifies the type for the midpoint value of a region map chart series.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.gradientMidpointType: "Percent" | ChartGradientStyleType | "ExtremeValue" | "Number"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMidpointValue",
        description: "Specifies the midpoint value of a region map chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.gradientMidpointValue: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMinimumColor",
        description: "Specifies the color for the minimum value of a region map chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.gradientMinimumColor: string",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMinimumType",
        description: "Specifies the type for the minimum value of a region map chart series.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.gradientMinimumType: "Percent" | ChartGradientStyleType | "ExtremeValue" | "Number"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientMinimumValue",
        description: "Specifies the minimum value of a region map chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.gradientMinimumValue: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.gradientStyle",
        description: "Specifies the series gradient style of a region map chart.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.gradientStyle: ChartGradientStyle | "TwoPhaseColor" | "ThreePhaseColor"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.hasDataLabels",
        description: "Specifies if the series has data labels.",
        kind: "Property",
        signature: "Excel.ChartSeries.hasDataLabels: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.invertColor",
        description: "Specifies the fill color for negative data points in a series.",
        kind: "Property",
        signature: "Excel.ChartSeries.invertColor: string",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.invertIfNegative",
        description:
          "True if Excel inverts the pattern in the item when it corresponds to a negative number.",
        kind: "Property",
        signature: "Excel.ChartSeries.invertIfNegative: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.mapOptions",
        description: "Encapsulates the options for a region map chart.",
        kind: "Property",
        signature: "Excel.ChartSeries.mapOptions: ChartMapOptions",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.markerBackgroundColor",
        description: "Specifies the marker background color of a chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.markerBackgroundColor: string",
        examples: ['series3.markerBackgroundColor = "purple";'],
      },
      {
        name: "Excel.ChartSeries.markerForegroundColor",
        description: "Specifies the marker foreground color of a chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.markerForegroundColor: string",
        examples: [
          'series0.markerForegroundColor = "black";',
          'series1.markerForegroundColor = "black";',
        ],
      },
      {
        name: "Excel.ChartSeries.markerSize",
        description: "Specifies the marker size of a chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.markerSize: number",
        examples: ["series2.markerSize = 12;"],
      },
      {
        name: "Excel.ChartSeries.markerStyle",
        description:
          "Specifies the marker style of a chart series. See `Excel.ChartMarkerStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.markerStyle: "Invalid" | Excel.ChartMarkerStyle | "Automatic" | "None" | "Square" | "Diamond" | "Triangle" | "X" | "Star" | "Dot" | "Dash" | "Circle" | "Plus" | "Picture"',
        examples: [
          'series0.markerStyle = "Dash";',
          'series1.markerStyle = "Star";',
          'series2.markerStyle = "X";',
          'series3.markerStyle = "Triangle";',
        ],
      },
      {
        name: "Excel.ChartSeries.name",
        description:
          "Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.",
        kind: "Property",
        signature: "Excel.ChartSeries.name: string",
        examples: [
          'activeChart.series.getItemAt(0).name = "New Series Name";',
          "seriesCollection.items[0].name;",
        ],
      },
      {
        name: "Excel.ChartSeries.overlap",
        description:
          "Specifies how bars and columns are positioned. Can be a value between â€“100 and 100. Applies only to 2-D bar and 2-D column charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.overlap: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.parentLabelStrategy",
        description: "Specifies the series parent label strategy area for a treemap chart.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.parentLabelStrategy: "None" | ChartParentLabelStrategy | "Banner" | "Overlapping"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.plotOrder",
        description: "Specifies the plot order of a chart series within the chart group.",
        kind: "Property",
        signature: "Excel.ChartSeries.plotOrder: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.points",
        description: "Returns a collection of all points in the series.",
        kind: "Property",
        signature: "Excel.ChartSeries.points: Excel.ChartPointsCollection",
        examples: [
          "let pointsCollection = activeChart.series.getItemAt(0).points;",
          "const points = activeChart.series.getItemAt(0).points;",
          "const pointsCollection = activeChart.series.getItemAt(0).points;",
        ],
      },
      {
        name: "Excel.ChartSeries.secondPlotSize",
        description:
          "Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.",
        kind: "Property",
        signature: "Excel.ChartSeries.secondPlotSize: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.showConnectorLines",
        description: "Specifies whether connector lines are shown in waterfall charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.showConnectorLines: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.showLeaderLines",
        description:
          "Specifies whether leader lines are displayed for each data label in the series.",
        kind: "Property",
        signature: "Excel.ChartSeries.showLeaderLines: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.showShadow",
        description: "Specifies if the series has a shadow.",
        kind: "Property",
        signature: "Excel.ChartSeries.showShadow: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.smooth",
        description:
          "Specifies if the series is smooth. Only applicable to line and scatter charts.",
        kind: "Property",
        signature: "Excel.ChartSeries.smooth: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.splitType",
        description:
          "Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.",
        kind: "Property",
        signature:
          'Excel.ChartSeries.splitType: ChartSplitType | "SplitByPosition" | "SplitByValue" | "SplitByPercentValue" | "SplitByCustomSplit"',
        examples: [],
      },
      {
        name: "Excel.ChartSeries.splitValue",
        description:
          "Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.",
        kind: "Property",
        signature: "Excel.ChartSeries.splitValue: number",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.trendlines",
        description: "The collection of trendlines in the series.",
        kind: "Property",
        signature: "Excel.ChartSeries.trendlines: Excel.ChartTrendlineCollection",
        examples: [
          'seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;',
          'series.trendlines.getItem(0).type = "Linear";',
          "let trendline = seriesCollection.getItemAt(0).trendlines.getItem(0);",
          'seriesCollection.getItemAt(0).trendlines.add("Linear");',
        ],
      },
      {
        name: "Excel.ChartSeries.varyByCategories",
        description:
          "True if Excel assigns a different color or pattern to each data marker. The chart must contain only one series.",
        kind: "Property",
        signature: "Excel.ChartSeries.varyByCategories: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.xErrorBars",
        description: "Represents the error bar object of a chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.xErrorBars: ChartErrorBars",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.yErrorBars",
        description: "Represents the error bar object of a chart series.",
        kind: "Property",
        signature: "Excel.ChartSeries.yErrorBars: ChartErrorBars",
        examples: [],
      },
      {
        name: "Excel.ChartSeries.delete",
        description: "Deletes the chart series.",
        kind: "Method",
        signature: "Excel.ChartSeries.delete() => void",
        examples: ["series.delete();", "bubbleChart.series.getItemAt(0).delete();"],
      },
      {
        name: "Excel.ChartSeries.getDimensionDataSourceString",
        description:
          "Gets the string representation of the data source of the chart series.The string representation could be information such as a cell address.",
        kind: "Method",
        signature:
          "Excel.ChartSeries.getDimensionDataSourceString(dimension: Excel.ChartSeriesDimension): OfficeExtension.ClientResult<string>",
        examples: ['const dataSourceString = series.getDimensionDataSourceString("Values");'],
      },
      {
        name: "Excel.ChartSeries.getDimensionDataSourceType",
        description: "Gets the data source type of the chart series.",
        kind: "Method",
        signature:
          "Excel.ChartSeries.getDimensionDataSourceType(dimension: Excel.ChartSeriesDimension): OfficeExtension.ClientResult<Excel.ChartDataSourceType>",
        examples: ['const dataSourceType = series.getDimensionDataSourceType("Values");'],
      },
      {
        name: "Excel.ChartSeries.getDimensionValues",
        description:
          "Gets the values from a single dimension of the chart series. These could be either category values or data values, depending on the dimension specified and how the data is mapped for the chart series.",
        kind: "Method",
        signature:
          "Excel.ChartSeries.getDimensionValues(dimension: Excel.ChartSeriesDimension): OfficeExtension.ClientResult<string[]>",
        examples: [
          "const bubbleSize = firstSeries.getDimensionValues(Excel.ChartSeriesDimension.bubbleSizes);",
          "const xValues = firstSeries.getDimensionValues(Excel.ChartSeriesDimension.xvalues);",
          "const yValues = firstSeries.getDimensionValues(Excel.ChartSeriesDimension.yvalues);",
          "const category = firstSeries.getDimensionValues(Excel.ChartSeriesDimension.categories);",
        ],
      },
      {
        name: "Excel.ChartSeries.setBubbleSizes",
        description: "Sets the bubble sizes for a chart series. Only works for bubble charts.",
        kind: "Method",
        signature: "Excel.ChartSeries.setBubbleSizes(sourceData: Excel.Range) => void",
        examples: ["newSeries.setBubbleSizes(dataRange.getCell(i, 3));"],
      },
      {
        name: "Excel.ChartSeries.setValues",
        description:
          "Sets the values for a chart series. For scatter charts, it refers to y-axis values.",
        kind: "Method",
        signature: "Excel.ChartSeries.setValues(sourceData: Excel.Range) => void",
        examples: [
          "newSeries.setValues(dataRange);",
          "newSeries.setValues(dataRange.getCell(i, 2));",
          "newSeries.setValues(rangeSelection);",
        ],
      },
      {
        name: "Excel.ChartSeries.setXAxisValues",
        description:
          "Sets the values of the x-axis for a chart series. Only works for scatter charts.",
        kind: "Method",
        signature: "Excel.ChartSeries.setXAxisValues(sourceData: Excel.Range) => void",
        examples: [
          "newSeries.setXAxisValues(dataRange.getCell(i, 1));",
          "newSeries.setXAxisValues(xRangeSelection);",
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartSeriesCollection",
    apiList: [
      {
        name: "Excel.ChartSeriesCollection.count",
        description: "Returns the number of series in the collection.",
        kind: "Property",
        signature: "Excel.ChartSeriesCollection.count: number",
        examples: ['"series: Count= " + seriesCollection.count;'],
      },
      {
        name: "Excel.ChartSeriesCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ChartSeriesCollection.items: Excel.ChartSeries[]",
        examples: ["seriesCollection.items[0].name;"],
      },
      {
        name: "Excel.ChartSeriesCollection.add",
        description:
          "Add a new series to the collection. The new added series is not visible until values, x-axis values, or bubble sizes for it are set (depending on chart type).",
        kind: "Method",
        signature:
          "Excel.ChartSeriesCollection.add(name?: string, index?: number) => Excel.ChartSeries",
        examples: [
          'let newSeries = activeChart.series.add("2016");',
          "const newSeries = bubbleChart.series.add(dataRange.values[i][0], i);",
          'let newSeries = activeChart.series.add("Qtr2");',
        ],
      },
      {
        name: "Excel.ChartSeriesCollection.getCount",
        description: "Returns the number of series in the collection.",
        kind: "Method",
        signature:
          "Excel.ChartSeriesCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.ChartSeriesCollection.getItemAt",
        description: "Retrieves a series based on its position in the collection.",
        kind: "Method",
        signature: "Excel.ChartSeriesCollection.getItemAt(index: number) => Excel.ChartSeries",
        examples: [
          'seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;',
          "let series = seriesCollection.getItemAt(0);",
          "let pointsCollection = activeChart.series.getItemAt(0).points;",
          "const points = activeChart.series.getItemAt(0).points;",
          "const pointsCollection = activeChart.series.getItemAt(0).points;",
          "const series = seriesCollection.getItemAt(0);",
          "const firstSeries = activeChart.series.getItemAt(0);",
          'activeChart.series.getItemAt(0).name = "New Series Name";',
          "let series0 = series.getItemAt(0);",
          "let series1 = series.getItemAt(1);",
          "let series2 = series.getItemAt(2);",
          "let series3 = series.getItemAt(3);",
          "bubbleChart.series.getItemAt(0).delete();",
          "let trendline = seriesCollection.getItemAt(0).trendlines.getItem(0);",
          'seriesCollection.getItemAt(0).trendlines.add("Linear");',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartSeriesFormat",
    apiList: [
      {
        name: "Excel.ChartSeriesFormat.fill",
        description:
          "Represents the fill format of a chart series, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartSeriesFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartSeriesFormat.line",
        description: "Represents line formatting.",
        kind: "Property",
        signature: "Excel.ChartSeriesFormat.line: ChartLineFormat",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartTitle",
    apiList: [
      {
        name: "Excel.ChartTitle.format",
        description:
          "Represents the formatting of a chart title, which includes fill and font formatting.",
        kind: "Property",
        signature: "Excel.ChartTitle.format: Excel.ChartTitleFormat",
        examples: [
          'title.format.font.name = "Calibri";',
          "title.format.font.size = 12;",
          'title.format.font.color = "#FF0000";',
          "title.format.font.italic = false;",
          "title.format.font.bold = true;",
          'title.format.font.underline = "None";',
        ],
      },
      {
        name: "Excel.ChartTitle.height",
        description:
          "Returns the height, in points, of the chart title. Value is `null` if the chart title is not visible.",
        kind: "Property",
        signature: "Excel.ChartTitle.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartTitle.horizontalAlignment",
        description: "Specifies the horizontal alignment for chart title.",
        kind: "Property",
        signature:
          'Excel.ChartTitle.horizontalAlignment: "Left" | "Center" | "Right" | "Justify" | "Distributed" | ChartTextHorizontalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartTitle.left",
        description:
          "Specifies the distance, in points, from the left edge of chart title to the left edge of chart area. Value is `null` if the chart title is not visible.",
        kind: "Property",
        signature: "Excel.ChartTitle.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartTitle.overlay",
        description: "Specifies if the chart title will overlay the chart.",
        kind: "Property",
        signature: "Excel.ChartTitle.overlay: boolean",
        examples: ["activeChart.title.overlay = true;"],
      },
      {
        name: "Excel.ChartTitle.position",
        description:
          "Represents the position of chart title. See `Excel.ChartTitlePosition` for details.",
        kind: "Property",
        signature:
          'Excel.ChartTitle.position: "Left" | "Right" | "Top" | "Bottom" | "Automatic" | ChartTitlePosition',
        examples: [],
      },
      {
        name: "Excel.ChartTitle.showShadow",
        description: "Represents a boolean value that determines if the chart title has a shadow.",
        kind: "Property",
        signature: "Excel.ChartTitle.showShadow: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartTitle.text",
        description: "Specifies the chart's title text.",
        kind: "Property",
        signature: "Excel.ChartTitle.text: string",
        examples: [
          'chart.title.text = "Sales Data";',
          'activeChart.title.text = "Sales Data by Year";',
          'chart.title.text = "Bicycle Parts Quarterly Sales";',
          'activeChart.title.text = "My Chart";',
        ],
      },
      {
        name: "Excel.ChartTitle.textOrientation",
        description:
          "Specifies the angle to which the text is oriented for the chart title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.",
        kind: "Property",
        signature: "Excel.ChartTitle.textOrientation: number",
        examples: ["title.textOrientation = -45;"],
      },
      {
        name: "Excel.ChartTitle.top",
        description:
          "Specifies the distance, in points, from the top edge of chart title to the top of chart area. Value is `null` if the chart title is not visible.",
        kind: "Property",
        signature: "Excel.ChartTitle.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartTitle.verticalAlignment",
        description:
          "Specifies the vertical alignment of chart title. See `Excel.ChartTextVerticalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.ChartTitle.verticalAlignment: "Center" | "Justify" | "Distributed" | "Top" | "Bottom" | ChartTextVerticalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartTitle.visible",
        description: "Specifies if the chart title is visibile.",
        kind: "Property",
        signature: "Excel.ChartTitle.visible: boolean",
        examples: ["activeChart.title.visible = true;"],
      },
      {
        name: "Excel.ChartTitle.width",
        description:
          "Specifies the width, in points, of the chart title. Value is `null` if the chart title is not visible.",
        kind: "Property",
        signature: "Excel.ChartTitle.width: number",
        examples: [],
      },
      {
        name: "Excel.ChartTitle.getSubstring",
        description: "Get the substring of a chart title. Line break '\\n' counts one character.",
        kind: "Method",
        signature:
          "Excel.ChartTitle.getSubstring(start: number, length: number) => Excel.ChartFormatString",
        examples: ['activeChart.title.getSubstring(0, 7).font.color = "Yellow";'],
      },
      {
        name: "Excel.ChartTitle.setFormula",
        description:
          "Sets a string value that represents the formula of chart title using A1-style notation.",
        kind: "Method",
        signature: "Excel.ChartTitle.setFormula => (formula: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartTitleFormat",
    apiList: [
      {
        name: "Excel.ChartTitleFormat.border",
        description:
          "Represents the border format of chart title, which includes color, linestyle, and weight.",
        kind: "Property",
        signature: "Excel.ChartTitleFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartTitleFormat.fill",
        description:
          "Represents the fill format of an object, which includes background formatting information.",
        kind: "Property",
        signature: "Excel.ChartTitleFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartTitleFormat.font",
        description:
          "Represents the font attributes (such as font name, font size, and color) for an object.",
        kind: "Property",
        signature: "Excel.ChartTitleFormat.font: Excel.ChartFont",
        examples: [
          'title.format.font.name = "Calibri";',
          "title.format.font.size = 12;",
          'title.format.font.color = "#FF0000";',
          "title.format.font.italic = false;",
          "title.format.font.bold = true;",
          'title.format.font.underline = "None";',
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartTrendline",
    apiList: [
      {
        name: "Excel.ChartTrendline.backwardPeriod",
        description: "Represents the number of periods that the trendline extends backward.",
        kind: "Property",
        signature: "Excel.ChartTrendline.backwardPeriod: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.format",
        description: "Represents the formatting of a chart trendline.",
        kind: "Property",
        signature: "Excel.ChartTrendline.format: Excel.ChartTrendlineFormat",
        examples: ["let line = trendline.format.line;"],
      },
      {
        name: "Excel.ChartTrendline.forwardPeriod",
        description: "Represents the number of periods that the trendline extends forward.",
        kind: "Property",
        signature: "Excel.ChartTrendline.forwardPeriod: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.intercept",
        description:
          "Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.",
        kind: "Property",
        signature: "Excel.ChartTrendline.intercept: any",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.label",
        description: "Represents the label of a chart trendline.",
        kind: "Property",
        signature: "Excel.ChartTrendline.label: ChartTrendlineLabel",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.movingAveragePeriod",
        description:
          "Represents the period of a chart trendline. Only applicable to trendlines with the type `MovingAverage`.",
        kind: "Property",
        signature: "Excel.ChartTrendline.movingAveragePeriod: number",
        examples: [
          'seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;',
        ],
      },
      {
        name: "Excel.ChartTrendline.name",
        description:
          "Represents the name of the trendline. Can be set to a string value, a `null` value represents automatic values. The returned value is always a string",
        kind: "Property",
        signature: "Excel.ChartTrendline.name: string",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.polynomialOrder",
        description:
          "Represents the order of a chart trendline. Only applicable to trendlines with the type `Polynomial`.",
        kind: "Property",
        signature: "Excel.ChartTrendline.polynomialOrder: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.showEquation",
        description: "True if the equation for the trendline is displayed on the chart.",
        kind: "Property",
        signature: "Excel.ChartTrendline.showEquation: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.showRSquared",
        description: "True if the r-squared value for the trendline is displayed on the chart.",
        kind: "Property",
        signature: "Excel.ChartTrendline.showRSquared: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartTrendline.type",
        description: "Represents the type of a chart trendline.",
        kind: "Property",
        signature:
          'Excel.ChartTrendline.type: Excel.ChartTrendlineType | "Linear" | "Exponential" | "Logarithmic" | "MovingAverage" | "Polynomial" | "Power"',
        examples: [
          'series.trendlines.getItem(0).type = "Linear";',
          '"The trendline type is:" + trendline.type;',
        ],
      },
      {
        name: "Excel.ChartTrendline.delete",
        description: "Delete the trendline object.",
        kind: "Method",
        signature: "Excel.ChartTrendline.delete => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartTrendlineCollection",
    apiList: [
      {
        name: "Excel.ChartTrendlineCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ChartTrendlineCollection.items: ChartTrendline[]",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineCollection.add",
        description: "Adds a new trendline to trendline collection.",
        kind: "Method",
        signature:
          "Excel.ChartTrendlineCollection.add(type?: Excel.ChartTrendlineType): Excel.ChartTrendline",
        examples: [
          'seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;',
          'seriesCollection.getItemAt(0).trendlines.add("Linear");',
        ],
      },
      {
        name: "Excel.ChartTrendlineCollection.getCount",
        description: "Returns the number of trendlines in the collection.",
        kind: "Method",
        signature:
          "Excel.ChartTrendlineCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineCollection.getItem",
        description:
          "Gets a trendline object by index, which is the insertion order in the items array.",
        kind: "Method",
        signature: "Excel.ChartTrendlineCollection.getItem(index: number) => Excel.ChartTrendline",
        examples: [
          'series.trendlines.getItem(0).type = "Linear";',
          "let trendline = seriesCollection.getItemAt(0).trendlines.getItem(0);",
        ],
      },
    ],
  },
  {
    objName: "Excel.ChartTrendlineFormat",
    apiList: [
      {
        name: "Excel.ChartTrendlineFormat.line",
        description: "Represents chart line formatting.",
        kind: "Property",
        signature: "Excel.ChartTrendlineFormat.line: Excel.ChartLineFormat",
        examples: ["let line = trendline.format.line;"],
      },
    ],
  },
  {
    objName: "Excel.ChartTrendlineLabel",
    apiList: [
      {
        name: "Excel.ChartTrendlineLabel.autoText",
        description:
          "Specifies if the trendline label automatically generates appropriate text based on context.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.autoText: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.format",
        description: "The format of the chart trendline label.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.format: ChartTrendlineLabelFormat",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.formula",
        description:
          "String value that represents the formula of the chart trendline label using A1-style notation.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.formula: string",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.height",
        description:
          "Returns the height, in points, of the chart trendline label. Value is `null` if the chart trendline label is not visible.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.height: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.horizontalAlignment",
        description:
          "Represents the horizontal alignment of the chart trendline label. See `Excel.ChartTextHorizontalAlignment` for details. This property is valid only when `TextOrientation` of a trendline label is -90, 90, or 180.",
        kind: "Property",
        signature:
          'Excel.ChartTrendlineLabel.horizontalAlignment: "Left" | "Center" | "Right" | "Justify" | "Distributed" | ChartTextHorizontalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.left",
        description:
          "Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area. Value is `null` if the chart trendline label is not visible.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.left: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.linkNumberFormat",
        description:
          "Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.linkNumberFormat: boolean",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.numberFormat",
        description: "String value that represents the format code for the trendline label.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.text",
        description: "String representing the text of the trendline label on a chart.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.text: string",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.textOrientation",
        description:
          "Represents the angle to which the text is oriented for the chart trendline label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.textOrientation: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.top",
        description:
          "Represents the distance, in points, from the top edge of the chart trendline label to the top of the chart area. Value is `null` if the chart trendline label is not visible.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.top: number",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.verticalAlignment",
        description:
          "Represents the vertical alignment of the chart trendline label. See `Excel.ChartTextVerticalAlignment` for details. This property is valid only when `TextOrientation` of a trendline label is 0.",
        kind: "Property",
        signature:
          'Excel.ChartTrendlineLabel.verticalAlignment: "Center" | "Justify" | "Distributed" | "Top" | "Bottom" | ChartTextVerticalAlignment',
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabel.width",
        description:
          "Returns the width, in points, of the chart trendline label. Value is `null` if the chart trendline label is not visible.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabel.width: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ChartTrendlineLabelFormat",
    apiList: [
      {
        name: "Excel.ChartTrendlineLabelFormat.border",
        description: "Specifies the border format, which includes color, linestyle, and weight.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabelFormat.border: ChartBorder",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabelFormat.fill",
        description: "Specifies the fill format of the current chart trendline label.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabelFormat.fill: ChartFill",
        examples: [],
      },
      {
        name: "Excel.ChartTrendlineLabelFormat.font",
        description:
          "Specifies the font attributes (such as font name, font size, and color) for a chart trendline label.",
        kind: "Property",
        signature: "Excel.ChartTrendlineLabelFormat.font: ChartFont",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ColorScaleConditionalFormat",
    apiList: [
      {
        name: "Excel.ColorScaleConditionalFormat.criteria",
        description:
          "The criteria of the color scale. Midpoint is optional when using a two point color scale.",
        kind: "Property",
        signature:
          "Excel.ColorScaleConditionalFormat.criteria: Excel.ConditionalColorScaleCriteria",
        examples: ["conditionalFormat.colorScale.criteria = criteria;"],
      },
      {
        name: "Excel.ColorScaleConditionalFormat.threeColorScale",
        description:
          "If `true`, the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).",
        kind: "Property",
        signature: "Excel.ColorScaleConditionalFormat.threeColorScale: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ColumnPropertiesLoadOptions",
    apiList: [
      {
        name: "Excel.ColumnPropertiesLoadOptions.address",
        description: "Specifies whether to load on the `address` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.address: boolean",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.addressLocal",
        description:
          "Specifies whether to load on the `addressLocal` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.addressLocal: boolean",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.columnHidden",
        description:
          "Specifies whether to load on the `columnHidden` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.columnHidden: boolean",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.columnIndex",
        description:
          "Specifies whether to load on the `columnIndex` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.columnIndex: boolean",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.format",
        description: "Specifies whether to load on the `format` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature:
          "Excel.ColumnPropertiesLoadOptions.format: CellPropertiesFormatLoadOptions & { columnWidth?: boolean; }",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.hidden",
        description: "Specifies whether to load on the `hidden` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.hidden: boolean",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.hyperlink",
        description:
          "Specifies whether to load on the `hyperlink` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.hyperlink: boolean",
        examples: [],
      },
      {
        name: "Excel.ColumnPropertiesLoadOptions.style",
        description: "Specifies whether to load on the `style` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.ColumnPropertiesLoadOptions.style: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Comment",
    apiList: [
      {
        name: "Excel.Comment.authorEmail",
        description: "Gets the email of the comment's author.",
        kind: "Property",
        signature: "Excel.Comment.authorEmail: string",
        examples: [
          "`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`;",
        ],
      },
      {
        name: "Excel.Comment.authorName",
        description: "Gets the name of the comment's author.",
        kind: "Property",
        signature: "Excel.Comment.authorName: string",
        examples: [
          "`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`;",
        ],
      },
      {
        name: "Excel.Comment.content",
        description: "The comment's content. The string is plain text.",
        kind: "Property",
        signature: "Excel.Comment.content: string",
        examples: ['comment.content = "PLEASE add headers here.";'],
      },
      {
        name: "Excel.Comment.contentType",
        description: "Gets the content type of the comment.",
        kind: "Property",
        signature: 'Excel.Comment.contentType: ContentType | "Plain" | "Mention"',
        examples: [],
      },
      {
        name: "Excel.Comment.creationDate",
        description:
          "Gets the creation time of the comment. Returns `null` if the comment was converted from a note, since the comment does not have a creation date.",
        kind: "Property",
        signature: "Excel.Comment.creationDate: Date",
        examples: [
          "`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`;",
        ],
      },
      {
        name: "Excel.Comment.id",
        description: "Specifies the comment identifier.",
        kind: "Property",
        signature: "Excel.Comment.id: string",
        examples: [],
      },
      {
        name: "Excel.Comment.mentions",
        description: "Gets the entities (e.g., people) that are mentioned in comments.",
        kind: "Property",
        signature: "Excel.Comment.mentions: CommentMention[]",
        examples: [],
      },
      {
        name: "Excel.Comment.replies",
        description: "Represents a collection of reply objects associated with the comment.",
        kind: "Property",
        signature: "Excel.Comment.replies: Excel.CommentReplyCollection",
        examples: [
          'comment.replies.add("Thanks for the reminder!");',
          "let reply = comment.replies.getItemAt(0);",
          "comment.replies.getItemAt(0).delete();",
          "let replyCount = comment.replies.getCount();",
          "let reply = comment.replies.getItemAt(replyCount.value - 1);",
          "const reply = comment.replies.getItemAt(0);",
          'comment.replies.add("Add content to this worksheet.");',
        ],
      },
      {
        name: "Excel.Comment.resolved",
        description:
          "The comment thread status. A value of `true` means that the comment thread is resolved.",
        kind: "Property",
        signature: "Excel.Comment.resolved: boolean",
        examples: [
          "workbook.comments.getItemAt(0).resolved = true;",
          "activeWorksheet.comments.getItemAt(0).resolved = true;",
        ],
      },
      {
        name: "Excel.Comment.richContent",
        description:
          "Gets the rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.",
        kind: "Property",
        signature: "Excel.Comment.richContent: string",
        examples: [],
      },
      {
        name: "Excel.Comment.assignTask",
        description:
          "Assigns the task attached to the comment to the given user as an assignee. If there is no task, one will be created.",
        kind: "Method",
        signature:
          "Excel.Comment.assignTask => (assignee: Excel.EmailIdentity) => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.Comment.delete",
        description: "Deletes the comment and all the connected replies.",
        kind: "Method",
        signature: "Excel.Comment.delete() => void",
        examples: [
          'workbook.comments.getItemByCell("MyWorksheet!A2").delete();',
          'workbook.comments.getItemByCell("Comments!A2").delete();',
        ],
      },
      {
        name: "Excel.Comment.getLocation",
        description: "Gets the cell where this comment is located.",
        kind: "Method",
        signature: "Excel.Comment.getLocation => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Comment.getTask",
        description:
          "Gets the task associated with this comment. If there is no task for the comment thread, an `ItemNotFound` exception is thrown.",
        kind: "Method",
        signature: "Excel.Comment.getTask => () => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.Comment.getTaskOrNullObject",
        description:
          "Gets the task associated with this comment. If there is no task for the comment thread, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Comment.getTaskOrNullObject => () => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.Comment.updateMentions",
        description:
          "Updates the comment content with a specially formatted string and a list of mentions.",
        kind: "Method",
        signature:
          "Excel.Comment.updateMentions => (contentWithMentions: Excel.CommentRichContent) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CommentCollection",
    apiList: [
      {
        name: "Excel.CommentCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.CommentCollection.items: Comment[]",
        examples: [],
      },
      {
        name: "Excel.CommentCollection.add",
        description:
          "Creates a new comment with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.",
        kind: "Method",
        signature:
          "Excel.CommentCollection.add(cellAddress: string | Excel.Range, content: string | Excel.CommentRichContent, contentType?: Excel.ContentType): Excel.Comment",
        examples: [
          'comments.add("MyWorksheet!A2", "TODO: add data.");',
          'workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);',
          'activeWorksheet.comments.add("A2", "TODO: add data.");',
          'activeWorksheet.comments.add("A1", commentBody, Excel.ContentType.mention);',
        ],
      },
      {
        name: "Excel.CommentCollection.getCount",
        description: "Gets the number of comments in the collection.",
        kind: "Method",
        signature: "Excel.CommentCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.CommentCollection.getItem",
        description: "Gets a comment from the collection based on its ID.",
        kind: "Method",
        signature: "Excel.CommentCollection.getItem => (commentId: string) => Excel.Comment",
        examples: [],
      },
      {
        name: "Excel.CommentCollection.getItemAt",
        description: "Gets a comment from the collection based on its position.",
        kind: "Method",
        signature: "Excel.CommentCollection.getItemAt(index: number) => Excel.Comment",
        examples: [
          "let comment = workbook.comments.getItemAt(0);",
          "workbook.comments.getItemAt(0).resolved = true;",
          "const comment = activeWorksheet.comments.getItemAt(0);",
          "activeWorksheet.comments.getItemAt(0).resolved = true;",
        ],
      },
      {
        name: "Excel.CommentCollection.getItemByCell",
        description: "Gets the comment from the specified cell.",
        kind: "Method",
        signature:
          "Excel.CommentCollection.getItemByCell(cellAddress: string | Excel.Range) => Excel.Comment",
        examples: [
          'workbook.comments.getItemByCell("MyWorksheet!A2").delete();',
          'let comment = workbook.comments.getItemByCell("MyWorksheet!A2");',
          'workbook.comments.getItemByCell("Comments!A2").delete();',
          'const comment = workbook.comments.getItemByCell("Comments!A2");',
        ],
      },
      {
        name: "Excel.CommentCollection.getItemByReplyId",
        description: "Gets the comment to which the given reply is connected.",
        kind: "Method",
        signature: "Excel.CommentCollection.getItemByReplyId => (replyId: string) => Excel.Comment",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CommentReply",
    apiList: [
      {
        name: "Excel.CommentReply.authorEmail",
        description: "Gets the email of the comment reply's author.",
        kind: "Property",
        signature: "Excel.CommentReply.authorEmail: string",
        examples: [
          "`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`;",
        ],
      },
      {
        name: "Excel.CommentReply.authorName",
        description: "Gets the name of the comment reply's author.",
        kind: "Property",
        signature: "Excel.CommentReply.authorName: string",
        examples: [
          "`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`;",
        ],
      },
      {
        name: "Excel.CommentReply.content",
        description: "The comment reply's content. The string is plain text.",
        kind: "Property",
        signature: "Excel.CommentReply.content: string",
        examples: ['reply.content = "Never mind";', 'reply.content += " Please!";'],
      },
      {
        name: "Excel.CommentReply.contentType",
        description: "The content type of the reply.",
        kind: "Property",
        signature: 'Excel.CommentReply.contentType: ContentType | "Plain" | "Mention"',
        examples: [],
      },
      {
        name: "Excel.CommentReply.creationDate",
        description: "Gets the creation time of the comment reply.",
        kind: "Property",
        signature: "Excel.CommentReply.creationDate: Date",
        examples: [
          "`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`;",
        ],
      },
      {
        name: "Excel.CommentReply.id",
        description: "Specifies the comment reply identifier.",
        kind: "Property",
        signature: "Excel.CommentReply.id: string",
        examples: [],
      },
      {
        name: "Excel.CommentReply.mentions",
        description: "The entities (e.g., people) that are mentioned in comments.",
        kind: "Property",
        signature: "Excel.CommentReply.mentions: CommentMention[]",
        examples: [],
      },
      {
        name: "Excel.CommentReply.resolved",
        description:
          "The comment reply status. A value of `true` means the reply is in the resolved state.",
        kind: "Property",
        signature: "Excel.CommentReply.resolved: boolean",
        examples: [],
      },
      {
        name: "Excel.CommentReply.richContent",
        description:
          "The rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.",
        kind: "Property",
        signature: "Excel.CommentReply.richContent: string",
        examples: [],
      },
      {
        name: "Excel.CommentReply.assignTask",
        description:
          "Assigns the task attached to the comment to the given user as the sole assignee. If there is no task, one will be created.",
        kind: "Method",
        signature:
          "Excel.CommentReply.assignTask => (assignee: Excel.EmailIdentity) => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.CommentReply.delete",
        description: "Deletes the comment reply.",
        kind: "Method",
        signature: "Excel.CommentReply.delete() => void",
        examples: ["comment.replies.getItemAt(0).delete();"],
      },
      {
        name: "Excel.CommentReply.getLocation",
        description: "Gets the cell where this comment reply is located.",
        kind: "Method",
        signature: "Excel.CommentReply.getLocation => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.CommentReply.getParentComment",
        description: "Gets the parent comment of this reply.",
        kind: "Method",
        signature: "Excel.CommentReply.getParentComment => () => Excel.Comment",
        examples: [],
      },
      {
        name: "Excel.CommentReply.getTask",
        description:
          "Gets the task associated with this comment reply's thread. If there is no task for the comment thread, an `ItemNotFound` exception is thrown.",
        kind: "Method",
        signature: "Excel.CommentReply.getTask => () => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.CommentReply.getTaskOrNullObject",
        description:
          "Gets the task associated with this comment reply's thread. If there is no task for the comment thread, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.CommentReply.getTaskOrNullObject => () => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.CommentReply.updateMentions",
        description:
          "Updates the comment content with a specially formatted string and a list of mentions.",
        kind: "Method",
        signature:
          "Excel.CommentReply.updateMentions => (contentWithMentions: Excel.CommentRichContent) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CommentReplyCollection",
    apiList: [
      {
        name: "Excel.CommentReplyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.CommentReplyCollection.items: CommentReply[]",
        examples: [],
      },
      {
        name: "Excel.CommentReplyCollection.add",
        description: "Creates a comment reply for a comment.",
        kind: "Method",
        signature:
          "Excel.CommentReplyCollection.add(content: string | Excel.CommentRichContent, contentType?: Excel.ContentType): Excel.CommentReply",
        examples: [
          'comment.replies.add("Thanks for the reminder!");',
          'comment.replies.add("Add content to this worksheet.");',
        ],
      },
      {
        name: "Excel.CommentReplyCollection.getCount",
        description: "Gets the number of comment replies in the collection.",
        kind: "Method",
        signature:
          "Excel.CommentReplyCollection.getCount() => OfficeExtension.ClientResult<number>",
        examples: ["let replyCount = comment.replies.getCount();"],
      },
      {
        name: "Excel.CommentReplyCollection.getItem",
        description: "Returns a comment reply identified by its ID.",
        kind: "Method",
        signature:
          "Excel.CommentReplyCollection.getItem => (commentReplyId: string) => Excel.CommentReply",
        examples: [],
      },
      {
        name: "Excel.CommentReplyCollection.getItemAt",
        description: "Gets a comment reply based on its position in the collection.",
        kind: "Method",
        signature: "Excel.CommentReplyCollection.getItemAt(index: number) => Excel.CommentReply",
        examples: [
          "let reply = comment.replies.getItemAt(0);",
          "comment.replies.getItemAt(0).delete();",
          "let reply = comment.replies.getItemAt(replyCount.value - 1);",
          "const reply = comment.replies.getItemAt(0);",
        ],
      },
    ],
  },
  {
    objName: "Excel.CommentRichContent",
    apiList: [
      {
        name: "Excel.CommentRichContent.mentions",
        description:
          "An array containing all the entities (e.g., people) mentioned within the comment.",
        kind: "Property",
        signature: "Excel.CommentRichContent.mentions: CommentMention[]",
        examples: [],
      },
      {
        name: "Excel.CommentRichContent.richContent",
        description:
          "Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an ID attribute of 0, and the second mentioned entity has an ID attribute of 1).",
        kind: "Property",
        signature: "Excel.CommentRichContent.richContent: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalCellValueRule",
    apiList: [
      {
        name: "Excel.ConditionalCellValueRule.formula1",
        description: "The formula, if required, on which to evaluate the conditional format rule.",
        kind: "Property",
        signature: "Excel.ConditionalCellValueRule.formula1: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalCellValueRule.formula2",
        description: "The formula, if required, on which to evaluate the conditional format rule.",
        kind: "Property",
        signature: "Excel.ConditionalCellValueRule.formula2: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalCellValueRule.operator",
        description: "The operator of the cell value conditional format.",
        kind: "Property",
        signature:
          'Excel.ConditionalCellValueRule.operator: "Between" | "GreaterThan" | "LessThan" | "NotBetween" | "EqualTo" | "NotEqualTo" | "Invalid" | "GreaterThanOrEqual" | ConditionalCellValueOperator | "LessThanOrEqual"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalColorScaleCriteria",
    apiList: [
      {
        name: "Excel.ConditionalColorScaleCriteria.maximum",
        description: "The maximum point of the color scale criterion.",
        kind: "Property",
        signature: "Excel.ConditionalColorScaleCriteria.maximum: ConditionalColorScaleCriterion",
        examples: [],
      },
      {
        name: "Excel.ConditionalColorScaleCriteria.midpoint",
        description:
          "The midpoint of the color scale criterion, if the color scale is a 3-color scale.",
        kind: "Property",
        signature: "Excel.ConditionalColorScaleCriteria.midpoint: ConditionalColorScaleCriterion",
        examples: [],
      },
      {
        name: "Excel.ConditionalColorScaleCriteria.minimum",
        description: "The minimum point of the color scale criterion.",
        kind: "Property",
        signature: "Excel.ConditionalColorScaleCriteria.minimum: ConditionalColorScaleCriterion",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalColorScaleCriterion",
    apiList: [
      {
        name: "Excel.ConditionalColorScaleCriterion.color",
        description:
          "HTML color code representation of the color scale color (e.g., #FF0000 represents Red).",
        kind: "Property",
        signature: "Excel.ConditionalColorScaleCriterion.color: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalColorScaleCriterion.formula",
        description: "A number, a formula, or `null` (if `type` is `lowestValue`).",
        kind: "Property",
        signature: "Excel.ConditionalColorScaleCriterion.formula: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalColorScaleCriterion.type",
        description: "What the criterion conditional formula should be based on.",
        kind: "Property",
        signature:
          'Excel.ConditionalColorScaleCriterion.type: "Percent" | "Invalid" | "Number" | "LowestValue" | "HighestValue" | "Formula" | "Percentile" | ConditionalFormatColorCriterionType',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalDataBarNegativeFormat",
    apiList: [
      {
        name: "Excel.ConditionalDataBarNegativeFormat.borderColor",
        description:
          'HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange"). Value is "" (an empty string) if no border is present or set.',
        kind: "Property",
        signature: "Excel.ConditionalDataBarNegativeFormat.borderColor: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalDataBarNegativeFormat.fillColor",
        description:
          'HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.ConditionalDataBarNegativeFormat.fillColor: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalDataBarNegativeFormat.matchPositiveBorderColor",
        description:
          "Specifies if the negative data bar has the same border color as the positive data bar.",
        kind: "Property",
        signature: "Excel.ConditionalDataBarNegativeFormat.matchPositiveBorderColor: boolean",
        examples: [],
      },
      {
        name: "Excel.ConditionalDataBarNegativeFormat.matchPositiveFillColor",
        description:
          "Specifies if the negative data bar has the same fill color as the positive data bar.",
        kind: "Property",
        signature: "Excel.ConditionalDataBarNegativeFormat.matchPositiveFillColor: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalDataBarPositiveFormat",
    apiList: [
      {
        name: "Excel.ConditionalDataBarPositiveFormat.borderColor",
        description:
          'HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange"). Value is "" (an empty string) if no border is present or set.',
        kind: "Property",
        signature: "Excel.ConditionalDataBarPositiveFormat.borderColor: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalDataBarPositiveFormat.fillColor",
        description:
          'HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.ConditionalDataBarPositiveFormat.fillColor: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalDataBarPositiveFormat.gradientFill",
        description: "Specifies if the data bar has a gradient.",
        kind: "Property",
        signature: "Excel.ConditionalDataBarPositiveFormat.gradientFill: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalDataBarRule",
    apiList: [
      {
        name: "Excel.ConditionalDataBarRule.formula",
        description: "The formula, if required, on which to evaluate the data bar rule.",
        kind: "Property",
        signature: "Excel.ConditionalDataBarRule.formula: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalDataBarRule.type",
        description: "The type of rule for the data bar.",
        kind: "Property",
        signature:
          'Excel.ConditionalDataBarRule.type: "Automatic" | "Percent" | "Invalid" | "Number" | ConditionalFormatRuleType | "LowestValue" | "HighestValue" | "Formula" | "Percentile"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalFormat",
    apiList: [
      {
        name: "Excel.ConditionalFormat.cellValue",
        description:
          "Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.cellValue: Excel.CellValueConditionalFormat",
        examples: [
          'conditionalFormat.cellValue.format.font.color = "red";',
          'conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };',
          'cellValueFormat.cellValue.format.font.color = "blue";',
          'cellValueFormat.cellValue.format.fill.color = "lightgreen";',
          'cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };',
        ],
      },
      {
        name: "Excel.ConditionalFormat.cellValueOrNullObject",
        description:
          "Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.cellValueOrNullObject: CellValueConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.colorScale",
        description:
          "Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.colorScale: Excel.ColorScaleConditionalFormat",
        examples: ["conditionalFormat.colorScale.criteria = criteria;"],
      },
      {
        name: "Excel.ConditionalFormat.colorScaleOrNullObject",
        description:
          "Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.colorScaleOrNullObject: ColorScaleConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.custom",
        description:
          "Returns the custom conditional format properties if the current conditional format is a custom type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.custom: Excel.CustomConditionalFormat",
        examples: [
          "conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT(\"RC[-1]\",0),TRUE)';",
          'conditionalFormat.custom.format.font.color = "green";',
          "conditionalFormat.custom.rule.formula = '=INDIRECT(\"E\"&ROW())>0.75';",
          'conditionalFormat.custom.format.fill.color = "green";',
        ],
      },
      {
        name: "Excel.ConditionalFormat.customOrNullObject",
        description:
          "Returns the custom conditional format properties if the current conditional format is a custom type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.customOrNullObject: Excel.CustomConditionalFormat",
        examples: ["const cfCustom = cf.customOrNullObject;"],
      },
      {
        name: "Excel.ConditionalFormat.dataBar",
        description:
          "Returns the data bar properties if the current conditional format is a data bar.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.dataBar: Excel.DataBarConditionalFormat",
        examples: [
          "conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;",
        ],
      },
      {
        name: "Excel.ConditionalFormat.dataBarOrNullObject",
        description:
          "Returns the data bar properties if the current conditional format is a data bar.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.dataBarOrNullObject: DataBarConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.iconSet",
        description:
          "Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.iconSet: Excel.IconSetConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.iconSetOrNullObject",
        description:
          "Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.iconSetOrNullObject: Excel.IconSetConditionalFormat",
        examples: [
          "conditionalFormat.iconSetOrNullObject.style = Excel.IconSet.fourTrafficLights;",
        ],
      },
      {
        name: "Excel.ConditionalFormat.id",
        description:
          "The priority of the conditional format in the current `ConditionalFormatCollection`.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.id: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.preset",
        description:
          "Returns the preset criteria conditional format. See `Excel.PresetCriteriaConditionalFormat` for more details.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.preset: Excel.PresetCriteriaConditionalFormat",
        examples: [
          'conditionalFormat.preset.format.font.color = "white";',
          'conditionalFormat.preset.format.font.color = "red";',
          'presetFormat.preset.format.font.color = "red";',
          "presetFormat.preset.format.font.bold = true;",
          "presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };",
          "conditionalFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage };",
        ],
      },
      {
        name: "Excel.ConditionalFormat.presetOrNullObject",
        description:
          "Returns the preset criteria conditional format. See `Excel.PresetCriteriaConditionalFormat` for more details.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.presetOrNullObject: PresetCriteriaConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.priority",
        description:
          "The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also changes other conditional formats' priorities, to allow for a contiguous priority order. Use a negative priority to begin from the back. Priorities greater than the bounds will get and set to the maximum (or minimum if negative) priority. Also note that if you change the priority, you have to re-fetch a new copy of the object at that new priority location if you want to make further changes to it.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.priority: number",
        examples: ["presetFormat.priority = 1;", "cellValueFormat.priority = 0;"],
      },
      {
        name: "Excel.ConditionalFormat.stopIfTrue",
        description:
          "If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell. Value is `null` on data bars, icon sets, and color scales as there's no concept of `StopIfTrue` for these.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.stopIfTrue: boolean",
        examples: ["cellValueFormat.stopIfTrue = true;"],
      },
      {
        name: "Excel.ConditionalFormat.textComparison",
        description:
          'Returns the specific text conditional format properties if the current conditional format is a text type. For example, to format cells matching the word "Text".',
        kind: "Property",
        signature: "Excel.ConditionalFormat.textComparison: Excel.TextConditionalFormat",
        examples: [
          'conditionalFormat.textComparison.format.font.color = "red";',
          'conditionalFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Delayed" };',
        ],
      },
      {
        name: "Excel.ConditionalFormat.textComparisonOrNullObject",
        description:
          'Returns the specific text conditional format properties if the current conditional format is a text type. For example, to format cells matching the word "Text".',
        kind: "Property",
        signature: "Excel.ConditionalFormat.textComparisonOrNullObject: TextConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.topBottom",
        description:
          "Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type. For example, to format the top 10% or bottom 10 items.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.topBottom: Excel.TopBottomConditionalFormat",
        examples: [
          'conditionalFormat.topBottom.format.fill.color = "green";',
          'conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems" };',
        ],
      },
      {
        name: "Excel.ConditionalFormat.topBottomOrNullObject",
        description:
          "Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type. For example, to format the top 10% or bottom 10 items.",
        kind: "Property",
        signature: "Excel.ConditionalFormat.topBottomOrNullObject: TopBottomConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.type",
        description: "A type of conditional format. Only one can be set at a time.",
        kind: "Property",
        signature:
          'Excel.ConditionalFormat.type: "Custom" | ConditionalFormatType | "DataBar" | "ColorScale" | "IconSet" | "TopBottom" | "PresetCriteria" | "ContainsText" | "CellValue"',
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToCellValue",
        description: "Change the conditional format rule type to cell value.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormat.changeRuleToCellValue => (properties: Excel.ConditionalCellValueRule) => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToColorScale",
        description: "Change the conditional format rule type to color scale.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.changeRuleToColorScale => () => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToContainsText",
        description: "Change the conditional format rule type to text comparison.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormat.changeRuleToContainsText => (properties: Excel.ConditionalTextComparisonRule) => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToCustom",
        description: "Change the conditional format rule type to custom.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.changeRuleToCustom => (formula: string) => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToDataBar",
        description: "Change the conditional format rule type to data bar.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.changeRuleToDataBar => () => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToIconSet",
        description: "Change the conditional format rule type to icon set.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.changeRuleToIconSet => () => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToPresetCriteria",
        description: "Change the conditional format rule type to preset criteria.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormat.changeRuleToPresetCriteria => (properties: Excel.ConditionalPresetCriteriaRule) => void",
        examples: [
          "conditionalFormat.changeRuleToPresetCriteria({\n    criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage,\n  });",
        ],
      },
      {
        name: "Excel.ConditionalFormat.changeRuleToTopBottom",
        description: "Change the conditional format rule type to top/bottom.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormat.changeRuleToTopBottom => (properties: Excel.ConditionalTopBottomRule) => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.delete",
        description: "Deletes this conditional format.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.getRange",
        description:
          "Returns the range the conditonal format is applied to. Throws an error if the conditional format is applied to multiple ranges.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.getRange => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.getRangeOrNullObject",
        description:
          "Returns the range to which the conditonal format is applied. If the conditional format is applied to multiple ranges, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.getRangeOrNullObject => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.getRanges",
        description:
          "Returns the `RangeAreas`, comprising one or more rectangular ranges, to which the conditonal format is applied.",
        kind: "Method",
        signature: "Excel.ConditionalFormat.getRanges => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormat.setRanges",
        description: "Set the ranges that the conditonal format rule is applied to.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormat.setRanges => (ranges: Range | RangeAreas | string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalFormatCollection",
    apiList: [
      {
        name: "Excel.ConditionalFormatCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ConditionalFormatCollection.items: ConditionalFormat[]",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormatCollection.add",
        description: "Adds a new conditional format to the collection at the first/top priority.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormatCollection.add(type: Excel.ConditionalFormatType): Excel.ConditionalFormat",
        examples: [
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);",
          "const presetFormat = temperatureDataRange.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);",
          "const cellValueFormat = temperatureDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);",
          "const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);",
          "const conditionalFormat = activeTable.getDataBodyRange().conditionalFormats.add(Excel.ConditionalFormatType.custom);",
        ],
      },
      {
        name: "Excel.ConditionalFormatCollection.clearAll",
        description: "Clears all conditional formats active on the current specified range.",
        kind: "Method",
        signature: "Excel.ConditionalFormatCollection.clearAll() => void",
        examples: ["range.conditionalFormats.clearAll();"],
      },
      {
        name: "Excel.ConditionalFormatCollection.getCount",
        description: "Returns the number of conditional formats in the workbook.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormatCollection.getCount() => OfficeExtension.ClientResult<number>",
        examples: ["const cfCount = range.conditionalFormats.getCount();"],
      },
      {
        name: "Excel.ConditionalFormatCollection.getItem",
        description: "Returns a conditional format for the given ID.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormatCollection.getItem => (id: string) => Excel.ConditionalFormat",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormatCollection.getItemAt",
        description: "Returns a conditional format at the given index.",
        kind: "Method",
        signature:
          "Excel.ConditionalFormatCollection.getItemAt(index: number) => Excel.ConditionalFormat",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalFormatRule",
    apiList: [
      {
        name: "Excel.ConditionalFormatRule.formula",
        description: "The formula, if required, on which to evaluate the conditional format rule.",
        kind: "Property",
        signature: "Excel.ConditionalFormatRule.formula: string",
        examples: [
          "conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT(\"RC[-1]\",0),TRUE)';",
          'cfCustom.rule.formula = "=ISBLANK(A1)";',
          "conditionalFormat.custom.rule.formula = '=INDIRECT(\"E\"&ROW())>0.75';",
        ],
      },
      {
        name: "Excel.ConditionalFormatRule.formulaLocal",
        description:
          "The formula, if required, on which to evaluate the conditional format rule in the user's language.",
        kind: "Property",
        signature: "Excel.ConditionalFormatRule.formulaLocal: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalFormatRule.formulaR1C1",
        description:
          "The formula, if required, on which to evaluate the conditional format rule in R1C1-style notation.",
        kind: "Property",
        signature: "Excel.ConditionalFormatRule.formulaR1C1: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalPresetCriteriaRule",
    apiList: [
      {
        name: "Excel.ConditionalPresetCriteriaRule.criterion",
        description: "The criterion of the conditional format.",
        kind: "Property",
        signature:
          'Excel.ConditionalPresetCriteriaRule.criterion: "Tomorrow" | "Today" | "Yesterday" | "NextWeek" | "ThisWeek" | "LastWeek" | "NextMonth" | "ThisMonth" | "LastMonth" | "Blanks" | "Errors" | "Invalid" | "AboveAverage" | "BelowAverage" | ... 13 more ... | "DuplicateValues"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalRangeBorder",
    apiList: [
      {
        name: "Excel.ConditionalRangeBorder.color",
        description:
          'HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.ConditionalRangeBorder.color: string",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorder.sideIndex",
        description:
          "Constant value that indicates the specific side of the border. See `Excel.ConditionalRangeBorderIndex` for details.",
        kind: "Property",
        signature:
          'Excel.ConditionalRangeBorder.sideIndex: "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | ConditionalRangeBorderIndex',
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorder.style",
        description:
          "One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ConditionalRangeBorder.style: "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | ConditionalRangeBorderLineStyle',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalRangeBorderCollection",
    apiList: [
      {
        name: "Excel.ConditionalRangeBorderCollection.bottom",
        description: "Gets the bottom border.",
        kind: "Property",
        signature: "Excel.ConditionalRangeBorderCollection.bottom: ConditionalRangeBorder",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.count",
        description: "Number of border objects in the collection.",
        kind: "Property",
        signature: "Excel.ConditionalRangeBorderCollection.count: number",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ConditionalRangeBorderCollection.items: ConditionalRangeBorder[]",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.left",
        description: "Gets the left border.",
        kind: "Property",
        signature: "Excel.ConditionalRangeBorderCollection.left: ConditionalRangeBorder",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.right",
        description: "Gets the right border.",
        kind: "Property",
        signature: "Excel.ConditionalRangeBorderCollection.right: ConditionalRangeBorder",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.top",
        description: "Gets the top border.",
        kind: "Property",
        signature: "Excel.ConditionalRangeBorderCollection.top: ConditionalRangeBorder",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.getItem",
        description: "Gets a border object using its name.",
        kind: "Method",
        signature:
          'Excel.ConditionalRangeBorderCollection.getItem => { (index: ConditionalRangeBorderIndex): ConditionalRangeBorder; (index: "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight"): ConditionalRangeBorder; (index: string): Excel.ConditionalRangeBorder; }',
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeBorderCollection.getItemAt",
        description: "Gets a border object using its index.",
        kind: "Method",
        signature:
          "Excel.ConditionalRangeBorderCollection.getItemAt => (index: number) => Excel.ConditionalRangeBorder",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalRangeFill",
    apiList: [
      {
        name: "Excel.ConditionalRangeFill.color",
        description:
          'HTML color code representing the color of the fill, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.ConditionalRangeFill.color: string",
        examples: [
          'conditionalFormat.topBottom.format.fill.color = "green";',
          'cellValueFormat.cellValue.format.fill.color = "lightgreen";',
          'cfCustom.format.fill.color = "#00FF00";',
          'conditionalFormat.custom.format.fill.color = "green";',
        ],
      },
      {
        name: "Excel.ConditionalRangeFill.clear",
        description: "Resets the fill.",
        kind: "Method",
        signature: "Excel.ConditionalRangeFill.clear => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalRangeFont",
    apiList: [
      {
        name: "Excel.ConditionalRangeFont.bold",
        description: "Specifies if the font is bold.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFont.bold: boolean",
        examples: ["presetFormat.preset.format.font.bold = true;"],
      },
      {
        name: "Excel.ConditionalRangeFont.color",
        description:
          "HTML color code representation of the text color (e.g., #FF0000 represents Red).",
        kind: "Property",
        signature: "Excel.ConditionalRangeFont.color: string",
        examples: [
          'conditionalFormat.cellValue.format.font.color = "red";',
          'conditionalFormat.custom.format.font.color = "green";',
          'conditionalFormat.preset.format.font.color = "white";',
          'conditionalFormat.textComparison.format.font.color = "red";',
          'conditionalFormat.preset.format.font.color = "red";',
          'presetFormat.preset.format.font.color = "red";',
          'cellValueFormat.cellValue.format.font.color = "blue";',
        ],
      },
      {
        name: "Excel.ConditionalRangeFont.italic",
        description: "Specifies if the font is italic.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFont.italic: boolean",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeFont.strikethrough",
        description: "Specifies the strikethrough status of the font.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFont.strikethrough: boolean",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeFont.underline",
        description:
          "The type of underline applied to the font. See `Excel.ConditionalRangeFontUnderlineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ConditionalRangeFont.underline: "Double" | "None" | "Single" | ConditionalRangeFontUnderlineStyle',
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeFont.clear",
        description: "Resets the font formats.",
        kind: "Method",
        signature: "Excel.ConditionalRangeFont.clear => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalRangeFormat",
    apiList: [
      {
        name: "Excel.ConditionalRangeFormat.borders",
        description:
          "Collection of border objects that apply to the overall conditional format range.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFormat.borders: ConditionalRangeBorderCollection",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeFormat.fill",
        description: "Returns the fill object defined on the overall conditional format range.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFormat.fill: Excel.ConditionalRangeFill",
        examples: [
          'conditionalFormat.topBottom.format.fill.color = "green";',
          'cellValueFormat.cellValue.format.fill.color = "lightgreen";',
          'cfCustom.format.fill.color = "#00FF00";',
          'conditionalFormat.custom.format.fill.color = "green";',
        ],
      },
      {
        name: "Excel.ConditionalRangeFormat.font",
        description: "Returns the font object defined on the overall conditional format range.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFormat.font: Excel.ConditionalRangeFont",
        examples: [
          'conditionalFormat.cellValue.format.font.color = "red";',
          'conditionalFormat.custom.format.font.color = "green";',
          'conditionalFormat.preset.format.font.color = "white";',
          'conditionalFormat.textComparison.format.font.color = "red";',
          'conditionalFormat.preset.format.font.color = "red";',
          'presetFormat.preset.format.font.color = "red";',
          "presetFormat.preset.format.font.bold = true;",
          'cellValueFormat.cellValue.format.font.color = "blue";',
        ],
      },
      {
        name: "Excel.ConditionalRangeFormat.numberFormat",
        description:
          "Represents Excel's number format code for the given range. For more information about Excel number formatting, see Number format codes. Cleared if `null` is passed in.",
        kind: "Property",
        signature: "Excel.ConditionalRangeFormat.numberFormat: any",
        examples: [],
      },
      {
        name: "Excel.ConditionalRangeFormat.clearFormat",
        description:
          "Remove the format properties from a conditional format rule. This creates a rule with no format settings.",
        kind: "Method",
        signature: "Excel.ConditionalRangeFormat.clearFormat => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalTextComparisonRule",
    apiList: [
      {
        name: "Excel.ConditionalTextComparisonRule.operator",
        description: "The operator of the text conditional format.",
        kind: "Property",
        signature:
          'Excel.ConditionalTextComparisonRule.operator: "BeginsWith" | "EndsWith" | "Contains" | "Invalid" | ConditionalTextOperator | "NotContains"',
        examples: [],
      },
      {
        name: "Excel.ConditionalTextComparisonRule.text",
        description: "The text value of the conditional format.",
        kind: "Property",
        signature: "Excel.ConditionalTextComparisonRule.text: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConditionalTopBottomRule",
    apiList: [
      {
        name: "Excel.ConditionalTopBottomRule.rank",
        description:
          "The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.",
        kind: "Property",
        signature: "Excel.ConditionalTopBottomRule.rank: number",
        examples: [],
      },
      {
        name: "Excel.ConditionalTopBottomRule.type",
        description: "Format values based on the top or bottom rank.",
        kind: "Property",
        signature:
          'Excel.ConditionalTopBottomRule.type: "Invalid" | "BottomItems" | "BottomPercent" | "TopItems" | "TopPercent" | ConditionalTopBottomCriterionType',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ConnectErrorCellValue",
    apiList: [
      {
        name: "Excel.ConnectErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.ConnectErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.ConnectErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.ConnectErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.ConnectErrorCellValue.errorSubType",
        description: "Represents the type of `ConnectErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.ConnectErrorCellValue.errorSubType: "Unknown" | ConnectErrorCellValueSubType | "ServiceError" | "ExternalLinks" | "ExternalLinksNonCloudLocation" | "DataTypeNoConnection" | ... 11 more ... | "GenericServerError"',
        examples: [],
      },
      {
        name: "Excel.ConnectErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.ConnectErrorCellValue.errorType: ErrorCellValueType.connect | "Connect"',
        examples: [],
      },
      {
        name: "Excel.ConnectErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.ConnectErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CultureInfo",
    apiList: [
      {
        name: "Excel.CultureInfo.datetimeFormat",
        description:
          "Defines the culturally appropriate format of displaying date and time. This is based on current system culture settings.",
        kind: "Property",
        signature: "Excel.CultureInfo.datetimeFormat: Excel.DatetimeFormatInfo",
        examples: [
          "const systemLongDatePattern = workbook.application.cultureInfo.datetimeFormat.longDatePattern;",
          "const systemShortDatePattern = workbook.application.cultureInfo.datetimeFormat.shortDatePattern;",
          "const systemDateSeparator = workbook.application.cultureInfo.datetimeFormat.dateSeparator;",
          "const systemLongTimePattern = workbook.application.cultureInfo.datetimeFormat.longTimePattern;",
          "const systemTimeSeparator = workbook.application.cultureInfo.datetimeFormat.timeSeparator;",
        ],
      },
      {
        name: "Excel.CultureInfo.name",
        description:
          'Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us"). This is based on current system settings.',
        kind: "Property",
        signature: "Excel.CultureInfo.name: string",
        examples: [],
      },
      {
        name: "Excel.CultureInfo.numberFormat",
        description:
          "Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.",
        kind: "Property",
        signature: "Excel.CultureInfo.numberFormat: Excel.NumberFormatInfo",
        examples: [
          "const systemDecimalSeparator = workbook.application.cultureInfo.numberFormat.numberDecimalSeparator;",
          "const systemThousandsSeparator = workbook.application.cultureInfo.numberFormat.numberGroupSeparator;",
        ],
      },
    ],
  },
  {
    objName: "Excel.CustomConditionalFormat",
    apiList: [
      {
        name: "Excel.CustomConditionalFormat.format",
        description:
          "Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.",
        kind: "Property",
        signature: "Excel.CustomConditionalFormat.format: Excel.ConditionalRangeFormat",
        examples: [
          'conditionalFormat.custom.format.font.color = "green";',
          'cfCustom.format.fill.color = "#00FF00";',
          'conditionalFormat.custom.format.fill.color = "green";',
        ],
      },
      {
        name: "Excel.CustomConditionalFormat.rule",
        description: "Specifies the `Rule` object on this conditional format.",
        kind: "Property",
        signature: "Excel.CustomConditionalFormat.rule: Excel.ConditionalFormatRule",
        examples: [
          "conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT(\"RC[-1]\",0),TRUE)';",
          'cfCustom.rule.formula = "=ISBLANK(A1)";',
          "conditionalFormat.custom.rule.formula = '=INDIRECT(\"E\"&ROW())>0.75';",
        ],
      },
    ],
  },
  {
    objName: "Excel.CustomDataValidation",
    apiList: [
      {
        name: "Excel.CustomDataValidation.formula",
        description:
          "A custom data validation formula. This creates special input rules, such as preventing duplicates, or limiting the total in a range of cells.",
        kind: "Property",
        signature: "Excel.CustomDataValidation.formula: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CustomProperty",
    apiList: [
      {
        name: "Excel.CustomProperty.key",
        description:
          "The key of the custom property. The key is limited to 255 characters outside of Excel on the web (larger keys are automatically trimmed to 255 characters on other platforms).",
        kind: "Property",
        signature: "Excel.CustomProperty.key: string",
        examples: [],
      },
      {
        name: "Excel.CustomProperty.type",
        description: "The type of the value used for the custom property.",
        kind: "Property",
        signature:
          'Excel.CustomProperty.type: "Boolean" | "String" | "Date" | "Number" | DocumentPropertyType | "Float"',
        examples: [],
      },
      {
        name: "Excel.CustomProperty.value",
        description:
          "The value of the custom property. The value is limited to 255 characters outside of Excel on the web (larger values are automatically trimmed to 255 characters on other platforms).",
        kind: "Property",
        signature: "Excel.CustomProperty.value: any",
        examples: [],
      },
      {
        name: "Excel.CustomProperty.delete",
        description: "Deletes the custom property.",
        kind: "Method",
        signature: "Excel.CustomProperty.delete => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.CustomPropertyCollection",
    apiList: [
      {
        name: "Excel.CustomPropertyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.CustomPropertyCollection.items: CustomProperty[]",
        examples: [],
      },
      {
        name: "Excel.CustomPropertyCollection.add",
        description: "Creates a new or sets an existing custom property.",
        kind: "Method",
        signature:
          "Excel.CustomPropertyCollection.add => (key: string, value: any) => Excel.CustomProperty",
        examples: [],
      },
      {
        name: "Excel.CustomPropertyCollection.deleteAll",
        description: "Deletes all custom properties in this collection.",
        kind: "Method",
        signature: "Excel.CustomPropertyCollection.deleteAll => () => void",
        examples: [],
      },
      {
        name: "Excel.CustomPropertyCollection.getCount",
        description: "Gets the count of custom properties.",
        kind: "Method",
        signature:
          "Excel.CustomPropertyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.CustomPropertyCollection.getItem",
        description:
          "Gets a custom property object by its key, which is case-insensitive. Throws an error if the custom property does not exist.",
        kind: "Method",
        signature:
          "Excel.CustomPropertyCollection.getItem => (key: string) => Excel.CustomProperty",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataBarConditionalFormat",
    apiList: [
      {
        name: "Excel.DataBarConditionalFormat.axisColor",
        description:
          'HTML color code representing the color of the Axis line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange"). Value is "" (an empty string) if no axis is present or set.',
        kind: "Property",
        signature: "Excel.DataBarConditionalFormat.axisColor: string",
        examples: [],
      },
      {
        name: "Excel.DataBarConditionalFormat.axisFormat",
        description: "Representation of how the axis is determined for an Excel data bar.",
        kind: "Property",
        signature:
          'Excel.DataBarConditionalFormat.axisFormat: "None" | "Automatic" | ConditionalDataBarAxisFormat | "CellMidPoint"',
        examples: [],
      },
      {
        name: "Excel.DataBarConditionalFormat.barDirection",
        description: "Specifies the direction that the data bar graphic should be based on.",
        kind: "Property",
        signature:
          'Excel.DataBarConditionalFormat.barDirection: Excel.ConditionalDataBarDirection | "Context" | "LeftToRight" | "RightToLeft"',
        examples: [
          "conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;",
        ],
      },
      {
        name: "Excel.DataBarConditionalFormat.lowerBoundRule",
        description:
          "The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar. The `ConditionalDataBarRule` object must be set as a JSON object (use `x.lowerBoundRule = {...}` instead of `x.lowerBoundRule.formula = ...`).",
        kind: "Property",
        signature: "Excel.DataBarConditionalFormat.lowerBoundRule: ConditionalDataBarRule",
        examples: [],
      },
      {
        name: "Excel.DataBarConditionalFormat.negativeFormat",
        description: "Representation of all values to the left of the axis in an Excel data bar.",
        kind: "Property",
        signature:
          "Excel.DataBarConditionalFormat.negativeFormat: ConditionalDataBarNegativeFormat",
        examples: [],
      },
      {
        name: "Excel.DataBarConditionalFormat.positiveFormat",
        description: "Representation of all values to the right of the axis in an Excel data bar.",
        kind: "Property",
        signature:
          "Excel.DataBarConditionalFormat.positiveFormat: ConditionalDataBarPositiveFormat",
        examples: [],
      },
      {
        name: "Excel.DataBarConditionalFormat.showDataBarOnly",
        description: "If `true`, hides the values from the cells where the data bar is applied.",
        kind: "Property",
        signature: "Excel.DataBarConditionalFormat.showDataBarOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.DataBarConditionalFormat.upperBoundRule",
        description:
          "The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar. The `ConditionalDataBarRule` object must be set as a JSON object (use `x.upperBoundRule = {...}` instead of `x.upperBoundRule.formula = ...`).",
        kind: "Property",
        signature: "Excel.DataBarConditionalFormat.upperBoundRule: ConditionalDataBarRule",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataControllerClient",
    apiList: [
      {
        name: "Excel.DataControllerClient.addField",
        description: "Add a field to a well.",
        kind: "Method",
        signature:
          "Excel.DataControllerClient.addField => (wellId: number, fieldId: number, position: number) => void",
        examples: [],
      },
      {
        name: "Excel.DataControllerClient.getAssociatedFields",
        description:
          "Gets an array of JSON objects representing the fields associated with the specified well ID. The objects in the array have an ID (number) and name (string).",
        kind: "Method",
        signature:
          "Excel.DataControllerClient.getAssociatedFields => (wellId: number) => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.DataControllerClient.getAvailableFields",
        description:
          "Gets an array of JSON objects representing the fields that may be associated with the well ID. The objects in the array have an ID (number) and name (string).",
        kind: "Method",
        signature:
          "Excel.DataControllerClient.getAvailableFields => (wellId: number) => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.DataControllerClient.getWells",
        description:
          "Gets an array of JSON objects representing this visual's wells. The objects in the array have an ID (number) and name (string).",
        kind: "Method",
        signature:
          "Excel.DataControllerClient.getWells => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.DataControllerClient.moveField",
        description: "Move a field from one position to another in a well.",
        kind: "Method",
        signature:
          "Excel.DataControllerClient.moveField => (wellId: number, fromPosition: number, toPosition: number) => void",
        examples: [],
      },
      {
        name: "Excel.DataControllerClient.removeField",
        description: "Remove a field from a well.",
        kind: "Method",
        signature:
          "Excel.DataControllerClient.removeField => (wellId: number, position: number) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataPivotHierarchy",
    apiList: [
      {
        name: "Excel.DataPivotHierarchy.field",
        description: "Returns the PivotFields associated with the DataPivotHierarchy.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchy.field: PivotField",
        examples: [],
      },
      {
        name: "Excel.DataPivotHierarchy.id",
        description: "ID of the DataPivotHierarchy.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchy.id: string",
        examples: [],
      },
      {
        name: "Excel.DataPivotHierarchy.name",
        description: "Name of the DataPivotHierarchy.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchy.name: string",
        examples: [
          'farmDataHierarchy.name = "Percentage of Total Farm Sales";',
          'farmDataHierarchy.name = "Difference from A Farms";',
          'dataHierarchies.items[0].name = "Farm Sales";',
          'dataHierarchies.items[1].name = "Wholesale";',
        ],
      },
      {
        name: "Excel.DataPivotHierarchy.numberFormat",
        description: "Number format of the DataPivotHierarchy.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchy.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.DataPivotHierarchy.position",
        description: "Position of the DataPivotHierarchy.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchy.position: number",
        examples: [],
      },
      {
        name: "Excel.DataPivotHierarchy.showAs",
        description: "Specifies if the data should be shown as a specific summary calculation.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchy.showAs: Excel.ShowAsRule",
        examples: [
          "let farmShowAs = farmDataHierarchy.showAs;",
          "farmDataHierarchy.showAs = farmShowAs;",
          "let wholesaleShowAs = wholesaleDataHierarchy.showAs;",
          "wholesaleDataHierarchy.showAs = wholesaleShowAs;",
        ],
      },
      {
        name: "Excel.DataPivotHierarchy.summarizeBy",
        description: "Specifies if all items of the DataPivotHierarchy are shown.",
        kind: "Property",
        signature:
          'Excel.DataPivotHierarchy.summarizeBy: Excel.AggregationFunction | "Unknown" | "Automatic" | "Sum" | "Count" | "Average" | "Max" | "Min" | "Product" | "CountNumbers" | "StandardDeviation" | "StandardDeviationP" | "Variance" | "VarianceP"',
        examples: [
          "pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;",
          "pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;",
        ],
      },
      {
        name: "Excel.DataPivotHierarchy.setToDefault",
        description: "Reset the DataPivotHierarchy back to its default values.",
        kind: "Method",
        signature: "Excel.DataPivotHierarchy.setToDefault => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataPivotHierarchyCollection",
    apiList: [
      {
        name: "Excel.DataPivotHierarchyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.DataPivotHierarchyCollection.items: Excel.DataPivotHierarchy[]",
        examples: [
          "pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;",
          "pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;",
          'dataHierarchies.items[0].name = "Farm Sales";',
          'dataHierarchies.items[1].name = "Wholesale";',
        ],
      },
      {
        name: "Excel.DataPivotHierarchyCollection.add",
        description: "Adds the PivotHierarchy to the current axis.",
        kind: "Method",
        signature:
          "Excel.DataPivotHierarchyCollection.add(pivotHierarchy: Excel.PivotHierarchy) => Excel.DataPivotHierarchy",
        examples: [
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));',
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));',
        ],
      },
      {
        name: "Excel.DataPivotHierarchyCollection.getCount",
        description: "Gets the number of pivot hierarchies in the collection.",
        kind: "Method",
        signature:
          "Excel.DataPivotHierarchyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.DataPivotHierarchyCollection.getItem",
        description: "Gets a DataPivotHierarchy by its name or ID.",
        kind: "Method",
        signature:
          "Excel.DataPivotHierarchyCollection.getItem(name: string) => Excel.DataPivotHierarchy",
        examples: [
          'let farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");',
          'const farmDataHierarchy: Excel.DataPivotHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");',
        ],
      },
      {
        name: "Excel.DataPivotHierarchyCollection.remove",
        description: "Removes the PivotHierarchy from the current axis.",
        kind: "Method",
        signature:
          "Excel.DataPivotHierarchyCollection.remove => (DataPivotHierarchy: Excel.DataPivotHierarchy) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataValidation",
    apiList: [
      {
        name: "Excel.DataValidation.errorAlert",
        description: "Error alert when user enters invalid data.",
        kind: "Property",
        signature: "Excel.DataValidation.errorAlert: Excel.DataValidationErrorAlert",
        examples: [
          'range.dataValidation.errorAlert = {\n    message: "Sorry, only positive whole numbers are allowed",\n    showAlert: true,\n    style: Excel.DataValidationAlertStyle.stop,\n    title: "Negative or Decimal Number Entered",\n  };',
          'commentsRange.dataValidation.errorAlert = {\n    message: "It is redundant to include the baby name in the comment.",\n    showAlert: true,\n    style: "Information",\n    title: "Baby Name in Comment",\n  };',
          'rankingRange.dataValidation.errorAlert = {\n    message: "Sorry, only positive numbers are allowed",\n    showAlert: true,\n    style: "Stop",\n    title: "Negative Number Entered",\n  };',
        ],
      },
      {
        name: "Excel.DataValidation.ignoreBlanks",
        description:
          "Specifies if data validation will be performed on blank cells. Default is `true`.",
        kind: "Property",
        signature: "Excel.DataValidation.ignoreBlanks: boolean",
        examples: [],
      },
      {
        name: "Excel.DataValidation.prompt",
        description: "Prompt when users select a cell.",
        kind: "Property",
        signature: "Excel.DataValidation.prompt: Excel.DataValidationPrompt",
        examples: [
          'range.dataValidation.prompt = {\n    message: "Please enter a positive whole number.",\n    showPrompt: true,\n    title: "Positive Whole Numbers Only.",\n  };',
          'rankingRange.dataValidation.prompt = {\n    message: "Please enter a positive number.",\n    showPrompt: true,\n    title: "Positive numbers only.",\n  };',
        ],
      },
      {
        name: "Excel.DataValidation.rule",
        description:
          "Data validation rule that contains different type of data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidation.rule: Excel.DataValidationRule",
        examples: [
          "commentsRange.dataValidation.rule = redundantStringRule;",
          "rankingRange.dataValidation.rule = greaterThanZeroRule;",
          "nameRange.dataValidation.rule = approvedListRule;",
        ],
      },
      {
        name: "Excel.DataValidation.type",
        description: "Type of the data validation, see `Excel.DataValidationType` for details.",
        kind: "Property",
        signature:
          'Excel.DataValidation.type: "List" | "None" | DataValidationType | "WholeNumber" | "Decimal" | "Date" | "Time" | "TextLength" | "Custom" | "Inconsistent" | "MixedCriteria"',
        examples: [],
      },
      {
        name: "Excel.DataValidation.valid",
        description:
          "Represents if all cell values are valid according to the data validation rules. Returns `true` if all cell values are valid, or `false` if all cell values are invalid. Returns `null` if there are both valid and invalid cell values within the range.",
        kind: "Property",
        signature: "Excel.DataValidation.valid: boolean",
        examples: [],
      },
      {
        name: "Excel.DataValidation.clear",
        description: "Clears the data validation from the current range.",
        kind: "Method",
        signature: "Excel.DataValidation.clear() => void",
        examples: [
          "commentsRange.dataValidation.clear();",
          "rankingRange.dataValidation.clear();",
          "nameRange.dataValidation.clear();",
        ],
      },
      {
        name: "Excel.DataValidation.getInvalidCells",
        description:
          "Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an `ItemNotFound` error.",
        kind: "Method",
        signature: "Excel.DataValidation.getInvalidCells => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.DataValidation.getInvalidCellsOrNullObject",
        description:
          "Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return `null`.",
        kind: "Method",
        signature: "Excel.DataValidation.getInvalidCellsOrNullObject => () => Excel.RangeAreas",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataValidationErrorAlert",
    apiList: [
      {
        name: "Excel.DataValidationErrorAlert.message",
        description: "Represents the error alert message.",
        kind: "Property",
        signature: "Excel.DataValidationErrorAlert.message: string",
        examples: [],
      },
      {
        name: "Excel.DataValidationErrorAlert.showAlert",
        description:
          "Specifies whether to show an error alert dialog when a user enters invalid data. The default is `true`.",
        kind: "Property",
        signature: "Excel.DataValidationErrorAlert.showAlert: boolean",
        examples: [],
      },
      {
        name: "Excel.DataValidationErrorAlert.style",
        description:
          "The data validation alert type, please see `Excel.DataValidationAlertStyle` for details.",
        kind: "Property",
        signature:
          'Excel.DataValidationErrorAlert.style: "Warning" | DataValidationAlertStyle | "Stop" | "Information"',
        examples: [],
      },
      {
        name: "Excel.DataValidationErrorAlert.title",
        description: "Represents the error alert dialog title.",
        kind: "Property",
        signature: "Excel.DataValidationErrorAlert.title: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataValidationPrompt",
    apiList: [
      {
        name: "Excel.DataValidationPrompt.message",
        description: "Specifies the message of the prompt.",
        kind: "Property",
        signature: "Excel.DataValidationPrompt.message: string",
        examples: [],
      },
      {
        name: "Excel.DataValidationPrompt.showPrompt",
        description:
          "Specifies if a prompt is shown when a user selects a cell with data validation.",
        kind: "Property",
        signature: "Excel.DataValidationPrompt.showPrompt: boolean",
        examples: [],
      },
      {
        name: "Excel.DataValidationPrompt.title",
        description: "Specifies the title for the prompt.",
        kind: "Property",
        signature: "Excel.DataValidationPrompt.title: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DataValidationRule",
    apiList: [
      {
        name: "Excel.DataValidationRule.custom",
        description: "Custom data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.custom: CustomDataValidation",
        examples: [],
      },
      {
        name: "Excel.DataValidationRule.date",
        description: "Date data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.date: DateTimeDataValidation",
        examples: [],
      },
      {
        name: "Excel.DataValidationRule.decimal",
        description: "Decimal data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.decimal: BasicDataValidation",
        examples: [],
      },
      {
        name: "Excel.DataValidationRule.list",
        description: "List data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.list: ListDataValidation",
        examples: [],
      },
      {
        name: "Excel.DataValidationRule.textLength",
        description: "Text length data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.textLength: BasicDataValidation",
        examples: [],
      },
      {
        name: "Excel.DataValidationRule.time",
        description: "Time data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.time: DateTimeDataValidation",
        examples: [],
      },
      {
        name: "Excel.DataValidationRule.wholeNumber",
        description: "Whole number data validation criteria.",
        kind: "Property",
        signature: "Excel.DataValidationRule.wholeNumber: BasicDataValidation",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DateTimeDataValidation",
    apiList: [
      {
        name: "Excel.DateTimeDataValidation.formula1",
        description:
          'Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand. When setting the value, it can be passed in as a Date, a Range object, or a string formula (where the string is either a stringified date/time in ISO8601 format, a cell reference like "=A1", or a formula like "=MIN(A1, B1)"). When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.',
        kind: "Property",
        signature: "Excel.DateTimeDataValidation.formula1: string | Date | Range",
        examples: [],
      },
      {
        name: "Excel.DateTimeDataValidation.formula2",
        description:
          'With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan. When setting the value, it can be passed in as a Date, a Range object, or a string (where the string is either a stringified date/time in ISO8601 format, a cell reference like "=A1", or a formula like "=MIN(A1, B1)"). When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.',
        kind: "Property",
        signature: "Excel.DateTimeDataValidation.formula2: string | Date | Range",
        examples: [],
      },
      {
        name: "Excel.DateTimeDataValidation.operator",
        description: "The operator to use for validating the data.",
        kind: "Property",
        signature:
          'Excel.DateTimeDataValidation.operator: "Between" | "GreaterThan" | "GreaterThanOrEqualTo" | "LessThan" | "LessThanOrEqualTo" | DataValidationOperator | "NotBetween" | "EqualTo" | "NotEqualTo"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DatetimeFormatInfo",
    apiList: [
      {
        name: "Excel.DatetimeFormatInfo.dateSeparator",
        description:
          "Gets the string used as the date separator. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.DatetimeFormatInfo.dateSeparator: string",
        examples: [
          "const systemDateSeparator = workbook.application.cultureInfo.datetimeFormat.dateSeparator;",
        ],
      },
      {
        name: "Excel.DatetimeFormatInfo.longDatePattern",
        description:
          "Gets the format string for a long date value. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.DatetimeFormatInfo.longDatePattern: string",
        examples: [
          "const systemLongDatePattern = workbook.application.cultureInfo.datetimeFormat.longDatePattern;",
        ],
      },
      {
        name: "Excel.DatetimeFormatInfo.longTimePattern",
        description:
          "Gets the format string for a long time value. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.DatetimeFormatInfo.longTimePattern: string",
        examples: [
          "const systemLongTimePattern = workbook.application.cultureInfo.datetimeFormat.longTimePattern;",
        ],
      },
      {
        name: "Excel.DatetimeFormatInfo.shortDatePattern",
        description:
          "Gets the format string for a short date value. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.DatetimeFormatInfo.shortDatePattern: string",
        examples: [
          "const systemShortDatePattern = workbook.application.cultureInfo.datetimeFormat.shortDatePattern;",
        ],
      },
      {
        name: "Excel.DatetimeFormatInfo.timeSeparator",
        description:
          "Gets the string used as the time separator. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.DatetimeFormatInfo.timeSeparator: string",
        examples: [
          "const systemTimeSeparator = workbook.application.cultureInfo.datetimeFormat.timeSeparator;",
        ],
      },
    ],
  },
  {
    objName: "Excel.Div0ErrorCellValue",
    apiList: [
      {
        name: "Excel.Div0ErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.Div0ErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.Div0ErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.Div0ErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.Div0ErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.Div0ErrorCellValue.errorType: ErrorCellValueType.div0 | "Div0"',
        examples: [],
      },
      {
        name: "Excel.Div0ErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.Div0ErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DocumentProperties",
    apiList: [
      {
        name: "Excel.DocumentProperties.author",
        description: "The author of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.author: string",
        examples: ['docProperties.author = "Alex";'],
      },
      {
        name: "Excel.DocumentProperties.category",
        description: "The category of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.category: string",
        examples: ["docProperties.category = categoryValue;"],
      },
      {
        name: "Excel.DocumentProperties.comments",
        description: "The comments of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.comments: string",
        examples: ["docProperties.comments = commentsValue;"],
      },
      {
        name: "Excel.DocumentProperties.company",
        description: "The company of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.company: string",
        examples: ["docProperties.company = companyValue;"],
      },
      {
        name: "Excel.DocumentProperties.creationDate",
        description: "Gets the creation date of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.creationDate: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentProperties.custom",
        description: "Gets the collection of custom properties of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.custom: CustomPropertyCollection",
        examples: [],
      },
      {
        name: "Excel.DocumentProperties.keywords",
        description: "The keywords of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.keywords: string",
        examples: ["docProperties.keywords = keywordsValue;"],
      },
      {
        name: "Excel.DocumentProperties.lastAuthor",
        description: "Gets the last author of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.lastAuthor: string",
        examples: [],
      },
      {
        name: "Excel.DocumentProperties.manager",
        description: "The manager of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.manager: string",
        examples: ["docProperties.manager = managerValue;"],
      },
      {
        name: "Excel.DocumentProperties.revisionNumber",
        description: "Gets the revision number of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.revisionNumber: number",
        examples: [],
      },
      {
        name: "Excel.DocumentProperties.subject",
        description: "The subject of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.subject: string",
        examples: ["docProperties.subject = subjectValue;"],
      },
      {
        name: "Excel.DocumentProperties.title",
        description: "The title of the workbook.",
        kind: "Property",
        signature: "Excel.DocumentProperties.title: string",
        examples: ["docProperties.title = titleValue;"],
      },
    ],
  },
  {
    objName: "Excel.DocumentTask",
    apiList: [
      {
        name: "Excel.DocumentTask.assignees",
        description: "Returns a collection of assignees of the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.assignees: EmailIdentity[]",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.changes",
        description: "Gets the change records of the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.changes: DocumentTaskChangeCollection",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.comment",
        description: "Gets the comment associated with the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.comment: Comment",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.completedBy",
        description: "Gets the most recent user to have completed the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.completedBy: EmailIdentity",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.completedDateTime",
        description: "Gets the date and time that the task was completed. All dates are in UTC.",
        kind: "Property",
        signature: "Excel.DocumentTask.completedDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.createdBy",
        description: "Gets the user who created the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.createdBy: EmailIdentity",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.createdDateTime",
        description: "Gets the date and time that the task was created. All dates are in UTC.",
        kind: "Property",
        signature: "Excel.DocumentTask.createdDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.id",
        description: "Gets the ID of the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.id: string",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.percentComplete",
        description:
          "Specifies the completion percentage of the task. This is a value between 0 and 100, where 100 represents a completed task.",
        kind: "Property",
        signature: "Excel.DocumentTask.percentComplete: number",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.priority",
        description:
          "Specifies the priority of the task. This is a value between 0 and 10, where 0 represents the highest priority.",
        kind: "Property",
        signature: "Excel.DocumentTask.priority: number",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.startAndDueDateTime",
        description: "Gets or sets the date and time the task should start and is due.",
        kind: "Property",
        signature: "Excel.DocumentTask.startAndDueDateTime: DocumentTaskSchedule",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.title",
        description: "Specifies title of the task.",
        kind: "Property",
        signature: "Excel.DocumentTask.title: string",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.assign",
        description: "Adds the given user to the list of assignees attached to the task.",
        kind: "Method",
        signature: "Excel.DocumentTask.assign => (assignee: Excel.EmailIdentity) => void",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.unassign",
        description: "Removes the given user from the list of assignees attached to the task.",
        kind: "Method",
        signature: "Excel.DocumentTask.unassign => (assignee: Excel.EmailIdentity) => void",
        examples: [],
      },
      {
        name: "Excel.DocumentTask.unassignAll",
        description: "Removes all users from the list of assignees attached to the task.",
        kind: "Method",
        signature: "Excel.DocumentTask.unassignAll => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DocumentTaskChange",
    apiList: [
      {
        name: "Excel.DocumentTaskChange.assignee",
        description:
          "Represents the user assigned to the task for an `assign` change action, or the user unassigned from the task for an `unassign` change action.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.assignee: EmailIdentity",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.changedBy",
        description: "Represents the identity of the user who made the task change.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.changedBy: EmailIdentity",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.commentId",
        description:
          "Represents the ID of the comment or commentReply to which the task change is anchored.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.commentId: string",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.createdDateTime",
        description:
          "Represents creation date and time of the task change record. All dates are in UTC.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.createdDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.dueDateTime",
        description:
          "Represents the task's due date and time. It is used for the `setSchedule` change action. It is in UTC time zone.It can be set to `null` to remove the due date and time. It should be set together with `startDateTime` to avoid conflicts.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.dueDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.id",
        description: "The unique GUID of the task change.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.id: string",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.percentComplete",
        description:
          "Represents the task's completion percentage. It is used for the `setPercentComplete` change action. This is a value betwen 0 and 100, where 100 represents a completed task.Changing this value to 100 also completes the associated comment.Changing the completion from 100 to a lower value reactivates the associated comment.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.percentComplete: number",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.priority",
        description:
          "Represents the task's priority. It is used for the `setPriority` change action. This is a value between 0 and 10, with 5 being the default priority if not set, and where 0 represents the highest priority.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.priority: number",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.startDateTime",
        description:
          "Represents the task's start date and time. It is used for the `setSchedule` change action. It is in UTC time zone.It can be set to `null` to remove the start date and time. It should be set together with `dueDateTime` to avoid conflicts.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.startDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.title",
        description: "Represents the task's title. It is used for `setTitle` change action.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.title: string",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.type",
        description:
          "Represents the action type of the task change record. Some examples of action types are assign, undo, and setPriority.",
        kind: "Property",
        signature:
          'Excel.DocumentTaskChange.type: DocumentTaskChangeAction | "unknown" | "create" | "assign" | "unassign" | "unassignAll" | "setSchedule" | "setPercentComplete" | "setPriority" | "remove" | "restore" | "setTitle" | "undo"',
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.undoChangeId",
        description:
          "Represents the `DocumentTaskChange.id` property that was undone for the `undo` change action.",
        kind: "Property",
        signature: "Excel.DocumentTaskChange.undoChangeId: string",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChange.newObject",
        description: "Create a new instance of Excel.DocumentTaskChange object",
        kind: "Method",
        signature:
          "Excel.DocumentTaskChange.newObject => (context: OfficeExtension.ClientRequestContext) => Excel.DocumentTaskChange",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DocumentTaskChangeCollection",
    apiList: [
      {
        name: "Excel.DocumentTaskChangeCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.DocumentTaskChangeCollection.items: DocumentTaskChange[]",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChangeCollection.getCount",
        description: "Gets the number of change records in the collection for the task.",
        kind: "Method",
        signature:
          "Excel.DocumentTaskChangeCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskChangeCollection.getItemAt",
        description: "Gets a task change record by using its index in the collection.",
        kind: "Method",
        signature:
          "Excel.DocumentTaskChangeCollection.getItemAt => (index: number) => Excel.DocumentTaskChange",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DocumentTaskCollection",
    apiList: [
      {
        name: "Excel.DocumentTaskCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.DocumentTaskCollection.items: DocumentTask[]",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskCollection.getCount",
        description: "Gets the number of tasks in the collection.",
        kind: "Method",
        signature:
          "Excel.DocumentTaskCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskCollection.getItem",
        description: "Gets a task using its ID.",
        kind: "Method",
        signature: "Excel.DocumentTaskCollection.getItem => (key: string) => Excel.DocumentTask",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskCollection.getItemAt",
        description: "Gets a task by its index in the collection.",
        kind: "Method",
        signature:
          "Excel.DocumentTaskCollection.getItemAt => (index: number) => Excel.DocumentTask",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DocumentTaskSchedule",
    apiList: [
      {
        name: "Excel.DocumentTaskSchedule.dueDateTime",
        description: "Gets the date and time that the task is due. All dates are in UTC.",
        kind: "Property",
        signature: "Excel.DocumentTaskSchedule.dueDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.DocumentTaskSchedule.startDateTime",
        description: "Gets the date and time that the task should start. All dates are in UTC.",
        kind: "Property",
        signature: "Excel.DocumentTaskSchedule.startDateTime: Date",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.DoubleCellValue",
    apiList: [
      {
        name: "Excel.DoubleCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.DoubleCellValue.basicType: RangeValueType.double | "Double"',
        examples: [],
      },
      {
        name: "Excel.DoubleCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: "Excel.DoubleCellValue.basicValue: number",
        examples: [],
      },
      {
        name: "Excel.DoubleCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.DoubleCellValue.type: CellValueType.double | "Double"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EmailIdentity",
    apiList: [
      {
        name: "Excel.EmailIdentity.displayName",
        description: "Represents the user's display name.",
        kind: "Property",
        signature: "Excel.EmailIdentity.displayName: string",
        examples: [],
      },
      {
        name: "Excel.EmailIdentity.email",
        description: "Represents the user's email.",
        kind: "Property",
        signature: "Excel.EmailIdentity.email: string",
        examples: [],
      },
      {
        name: "Excel.EmailIdentity.id",
        description: "Represents the user's unique ID.",
        kind: "Property",
        signature: "Excel.EmailIdentity.id: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EmptyCellValue",
    apiList: [
      {
        name: "Excel.EmptyCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.EmptyCellValue.basicType: RangeValueType.empty | "Empty"',
        examples: [],
      },
      {
        name: "Excel.EmptyCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.EmptyCellValue.basicValue: ""',
        examples: [],
      },
      {
        name: "Excel.EmptyCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.EmptyCellValue.type: CellValueType.empty | "Empty"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EntityArrayCardLayout",
    apiList: [
      {
        name: "Excel.EntityArrayCardLayout.arrayProperty",
        description: "Represents name of the property that contains the array shown in the card.",
        kind: "Property",
        signature: "Excel.EntityArrayCardLayout.arrayProperty: string",
        examples: [],
      },
      {
        name: "Excel.EntityArrayCardLayout.columnsToReport",
        description:
          "Represents the count of columns which the card claims are in the array. A card may report a different number of columns than it actually has to display smaller amounts of preview data.",
        kind: "Property",
        signature: "Excel.EntityArrayCardLayout.columnsToReport: number",
        examples: [],
      },
      {
        name: "Excel.EntityArrayCardLayout.displayName",
        description:
          'Represents name of the property that contains the array shown in the card. Default value is "Array".',
        kind: "Property",
        signature: "Excel.EntityArrayCardLayout.displayName: string",
        examples: [],
      },
      {
        name: "Excel.EntityArrayCardLayout.firstRowIsHeader",
        description: "Represents whether the first row of the array is treated as a header.",
        kind: "Property",
        signature: "Excel.EntityArrayCardLayout.firstRowIsHeader: boolean",
        examples: [],
      },
      {
        name: "Excel.EntityArrayCardLayout.layout",
        description: "Represents the type of this layout.",
        kind: "Property",
        signature: 'Excel.EntityArrayCardLayout.layout: "Array" | EntityCardLayoutType.array',
        examples: [],
      },
      {
        name: "Excel.EntityArrayCardLayout.rowsToReport",
        description:
          "Represents the count of rows which the card claims are in the array. A card may report a different number of rows than it actually has to display smaller amounts of preview data.",
        kind: "Property",
        signature: "Excel.EntityArrayCardLayout.rowsToReport: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EntityCardLayout",
    apiList: [
      {
        name: "Excel.EntityCardLayout.layout",
        description: "Represents the type of this layout.",
        kind: "Property",
        signature: 'Excel.EntityCardLayout.layout: EntityCardLayoutType.entity | "Entity"',
        examples: [],
      },
      {
        name: "Excel.EntityCardLayout.mainImage",
        description: "Specifies a property which will be used as the main image of the card.",
        kind: "Property",
        signature: "Excel.EntityCardLayout.mainImage: CardLayoutPropertyReference",
        examples: [],
      },
      {
        name: "Excel.EntityCardLayout.sections",
        description: "Represents the sections of the card.",
        kind: "Property",
        signature: "Excel.EntityCardLayout.sections: CardLayoutSection[]",
        examples: [],
      },
      {
        name: "Excel.EntityCardLayout.subTitle",
        description:
          "Represents a specification of which property contains the subtitle of the card.",
        kind: "Property",
        signature: "Excel.EntityCardLayout.subTitle: CardLayoutPropertyReference",
        examples: [],
      },
      {
        name: "Excel.EntityCardLayout.title",
        description:
          "Represents the title of the card or the specification of which property contains the title of the card.",
        kind: "Property",
        signature: "Excel.EntityCardLayout.title: string | CardLayoutPropertyReference",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EntityCellValue",
    apiList: [
      {
        name: "Excel.EntityCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.EntityCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.EntityCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.layouts",
        description: "Represents layout information for views of this entity.",
        kind: "Property",
        signature: "Excel.EntityCellValue.layouts: EntityViewLayouts",
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.properties",
        description: "Represents the properties of this entity and their metadata.",
        kind: "Property",
        signature: "Excel.EntityCellValue.properties: { [key: string]: EntityPropertyType; }",
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.provider",
        description:
          "Represents information that describes the service that provided the data in this `EntityCellValue`. This information can be used for branding in entity cards.",
        kind: "Property",
        signature: "Excel.EntityCellValue.provider: CellValueProviderAttributes",
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.referencedValues",
        description:
          "Represents the cell values which are referenced within `EntityCellValue.properties`.",
        kind: "Property",
        signature: "Excel.EntityCellValue.referencedValues: ReferencedValue[]",
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.text",
        description: "Represents the text shown when a cell with this value is rendered.",
        kind: "Property",
        signature: "Excel.EntityCellValue.text: string",
        examples: [],
      },
      {
        name: "Excel.EntityCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature:
          'Excel.EntityCellValue.type: CellValueType.entity | ReferenceValueType.entity | "Entity"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EntityCompactLayout",
    apiList: [
      {
        name: "Excel.EntityCompactLayout.icon",
        description: "Specifies the name of the icon which is used to open the card.",
        kind: "Property",
        signature: "Excel.EntityCompactLayout.icon: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.EntityViewLayouts",
    apiList: [
      {
        name: "Excel.EntityViewLayouts.card",
        description:
          'Represents the layout of this entity in card view. If the `CardLayout` object does not have a layout property, it is assumed to be "Entity".',
        kind: "Property",
        signature: "Excel.EntityViewLayouts.card: CardLayout",
        examples: [],
      },
      {
        name: "Excel.EntityViewLayouts.compact",
        description:
          "Represents the layout used when there is limited space to represent the entity.",
        kind: "Property",
        signature: "Excel.EntityViewLayouts.compact: EntityCompactLayout",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ExternalErrorCellValue",
    apiList: [
      {
        name: "Excel.ExternalErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.ExternalErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.ExternalErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.ExternalErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.ExternalErrorCellValue.errorSubType",
        description: "Represents the type of `ExternalErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.ExternalErrorCellValue.errorSubType: "Unknown" | ExternalErrorCellValueSubType',
        examples: [],
      },
      {
        name: "Excel.ExternalErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.ExternalErrorCellValue.errorType: ErrorCellValueType.external | "External"',
        examples: [],
      },
      {
        name: "Excel.ExternalErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.ExternalErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FieldErrorCellValue",
    apiList: [
      {
        name: "Excel.FieldErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.FieldErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.FieldErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.FieldErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.FieldErrorCellValue.errorSubType",
        description: "Represents the type of `FieldErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.FieldErrorCellValue.errorSubType: "Unknown" | FieldErrorCellValueSubType | "WebImageMissingFilePart" | "DataProviderError"',
        examples: [],
      },
      {
        name: "Excel.FieldErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.FieldErrorCellValue.errorType: ErrorCellValueType.field | "Field"',
        examples: [],
      },
      {
        name: "Excel.FieldErrorCellValue.fieldName",
        description: "Represents the field which was not found by FIELDVALUE.",
        kind: "Property",
        signature: "Excel.FieldErrorCellValue.fieldName: string",
        examples: [],
      },
      {
        name: "Excel.FieldErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.FieldErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Filter",
    apiList: [
      {
        name: "Excel.Filter.criteria",
        description: "The currently applied filter on the given column.",
        kind: "Property",
        signature: "Excel.Filter.criteria: FilterCriteria",
        examples: [],
      },
      {
        name: "Excel.Filter.apply",
        description: "Apply the given filter criteria on the given column.",
        kind: "Method",
        signature: "Excel.Filter.apply(criteria: Excel.FilterCriteria) => void",
        examples: [
          'categoryFilter.apply({\n    filterOn: Excel.FilterOn.values,\n    values: ["Restaurant", "Groceries"],\n  });',
          "amountFilter.apply({\n    filterOn: Excel.FilterOn.dynamic,\n    dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage,\n  });",
          "filter.apply({\n    filterOn: Excel.FilterOn.dynamic,\n    dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage,\n  });",
          'filter.apply({\n    filterOn: Excel.FilterOn.values,\n    values: ["Restaurant", "Groceries"],\n  });',
        ],
      },
      {
        name: "Excel.Filter.applyBottomItemsFilter",
        description: 'Apply a "Bottom Item" filter to the column for the given number of elements.',
        kind: "Method",
        signature: "Excel.Filter.applyBottomItemsFilter => (count: number) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyBottomPercentFilter",
        description:
          'Apply a "Bottom Percent" filter to the column for the given percentage of elements.',
        kind: "Method",
        signature: "Excel.Filter.applyBottomPercentFilter => (percent: number) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyCellColorFilter",
        description: 'Apply a "Cell Color" filter to the column for the given color.',
        kind: "Method",
        signature: "Excel.Filter.applyCellColorFilter => (color: string) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyCustomFilter",
        description: 'Apply an "Icon" filter to the column for the given criteria strings.',
        kind: "Method",
        signature:
          'Excel.Filter.applyCustomFilter => { (criteria1: string, criteria2?: string, oper?: FilterOperator): void; (criteria1: string, criteria2?: string, oper?: "And" | "Or"): void; (criteria1: string, criteria2?: string, oper?: string): void; }',
        examples: [],
      },
      {
        name: "Excel.Filter.applyDynamicFilter",
        description: 'Apply a "Dynamic" filter to the column.',
        kind: "Method",
        signature:
          'Excel.Filter.applyDynamicFilter => { (criteria: DynamicFilterCriteria): void; (criteria: "Unknown" | "Tomorrow" | "Today" | "Yesterday" | "NextWeek" | "ThisWeek" | "LastWeek" | "NextMonth" | "ThisMonth" | ... 25 more ... | "BelowAverage"): void; (criteria: string): void; }',
        examples: [],
      },
      {
        name: "Excel.Filter.applyFontColorFilter",
        description: 'Apply a "Font Color" filter to the column for the given color.',
        kind: "Method",
        signature: "Excel.Filter.applyFontColorFilter => (color: string) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyIconFilter",
        description: 'Apply an "Icon" filter to the column for the given icon.',
        kind: "Method",
        signature: "Excel.Filter.applyIconFilter => (icon: Excel.Icon) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyTopItemsFilter",
        description: 'Apply a "Top Item" filter to the column for the given number of elements.',
        kind: "Method",
        signature: "Excel.Filter.applyTopItemsFilter => (count: number) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyTopPercentFilter",
        description:
          'Apply a "Top Percent" filter to the column for the given percentage of elements.',
        kind: "Method",
        signature: "Excel.Filter.applyTopPercentFilter => (percent: number) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.applyValuesFilter",
        description: 'Apply a "Values" filter to the column for the given values.',
        kind: "Method",
        signature:
          "Excel.Filter.applyValuesFilter => (values: Array<string | FilterDatetime>) => void",
        examples: [],
      },
      {
        name: "Excel.Filter.clear",
        description: "Clear the filter on the given column.",
        kind: "Method",
        signature: "Excel.Filter.clear => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FilterCriteria",
    apiList: [
      {
        name: "Excel.FilterCriteria.color",
        description:
          "The HTML color string used to filter cells. Used with `cellColor` and `fontColor` filtering.",
        kind: "Property",
        signature: "Excel.FilterCriteria.color: string",
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.criterion1",
        description:
          'The first criterion used to filter data. Used as an operator in the case of `custom` filtering. For example ">50" for numbers greater than 50, or "=*s" for values ending in "s". Used as a number in the case of top/bottom items/percents (e.g., "5" for the top 5 items if `filterOn` is set to `topItems`).',
        kind: "Property",
        signature: "Excel.FilterCriteria.criterion1: string",
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.criterion2",
        description:
          "The second criterion used to filter data. Only used as an operator in the case of `custom` filtering.",
        kind: "Property",
        signature: "Excel.FilterCriteria.criterion2: string",
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.dynamicCriteria",
        description:
          "The dynamic criteria from the `Excel.DynamicFilterCriteria` set to apply on this column. Used with `dynamic` filtering.",
        kind: "Property",
        signature:
          'Excel.FilterCriteria.dynamicCriteria: "Unknown" | "Tomorrow" | "Today" | "Yesterday" | "NextWeek" | "ThisWeek" | "LastWeek" | "NextMonth" | "ThisMonth" | "LastMonth" | "NextQuarter" | "ThisQuarter" | "LastQuarter" | ... 22 more ... | "BelowAverage"',
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.filterOn",
        description:
          "The property used by the filter to determine whether the values should stay visible.",
        kind: "Property",
        signature:
          'Excel.FilterCriteria.filterOn: "Values" | "Custom" | "CellColor" | "FontColor" | "Icon" | FilterOn | "BottomItems" | "BottomPercent" | "Dynamic" | "TopItems" | "TopPercent"',
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.icon",
        description: "The icon used to filter cells. Used with `icon` filtering.",
        kind: "Property",
        signature: "Excel.FilterCriteria.icon: Icon",
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.operator",
        description:
          "The operator used to combine criterion 1 and 2 when using `custom` filtering.",
        kind: "Property",
        signature: 'Excel.FilterCriteria.operator: FilterOperator | "And" | "Or"',
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.subField",
        description: "The property used by the filter to do a rich filter on rich values.",
        kind: "Property",
        signature: "Excel.FilterCriteria.subField: string",
        examples: [],
      },
      {
        name: "Excel.FilterCriteria.values",
        description: "The set of values to be used as part of `values` filtering.",
        kind: "Property",
        signature: "Excel.FilterCriteria.values: (string | FilterDatetime)[]",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FilterDatetime",
    apiList: [
      {
        name: "Excel.FilterDatetime.date",
        description: "The date in ISO8601 format used to filter data.",
        kind: "Property",
        signature: "Excel.FilterDatetime.date: string",
        examples: [],
      },
      {
        name: "Excel.FilterDatetime.specificity",
        description:
          'How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specificity is set to "month", the filter operation will keep all rows with a date in the month of April 2005.',
        kind: "Property",
        signature:
          'Excel.FilterDatetime.specificity: FilterDatetimeSpecificity | "Year" | "Month" | "Day" | "Hour" | "Minute" | "Second"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FilterPivotHierarchy",
    apiList: [
      {
        name: "Excel.FilterPivotHierarchy.enableMultipleFilterItems",
        description: "Determines whether to allow multiple filter items.",
        kind: "Property",
        signature: "Excel.FilterPivotHierarchy.enableMultipleFilterItems: boolean",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchy.fields",
        description: "Returns the PivotFields associated with the FilterPivotHierarchy.",
        kind: "Property",
        signature: "Excel.FilterPivotHierarchy.fields: Excel.PivotFieldCollection",
        examples: ['const filterField = classHierarchy.fields.getItem("Classification");'],
      },
      {
        name: "Excel.FilterPivotHierarchy.id",
        description: "ID of the FilterPivotHierarchy.",
        kind: "Property",
        signature: "Excel.FilterPivotHierarchy.id: string",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchy.name",
        description: "Name of the FilterPivotHierarchy.",
        kind: "Property",
        signature: "Excel.FilterPivotHierarchy.name: string",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchy.position",
        description: "Position of the FilterPivotHierarchy.",
        kind: "Property",
        signature: "Excel.FilterPivotHierarchy.position: number",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchy.setToDefault",
        description: "Reset the FilterPivotHierarchy back to its default values.",
        kind: "Method",
        signature: "Excel.FilterPivotHierarchy.setToDefault => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FilterPivotHierarchyCollection",
    apiList: [
      {
        name: "Excel.FilterPivotHierarchyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.FilterPivotHierarchyCollection.items: FilterPivotHierarchy[]",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchyCollection.add",
        description:
          "Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column, or filter axis, it will be removed from that location.",
        kind: "Method",
        signature:
          "Excel.FilterPivotHierarchyCollection.add(pivotHierarchy: Excel.PivotHierarchy) => Excel.FilterPivotHierarchy",
        examples: [
          'classHierarchy = pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
        ],
      },
      {
        name: "Excel.FilterPivotHierarchyCollection.getCount",
        description: "Gets the number of pivot hierarchies in the collection.",
        kind: "Method",
        signature:
          "Excel.FilterPivotHierarchyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchyCollection.getItem",
        description: "Gets a FilterPivotHierarchy by its name or ID.",
        kind: "Method",
        signature:
          "Excel.FilterPivotHierarchyCollection.getItem => (name: string) => Excel.FilterPivotHierarchy",
        examples: [],
      },
      {
        name: "Excel.FilterPivotHierarchyCollection.remove",
        description: "Removes the PivotHierarchy from the current axis.",
        kind: "Method",
        signature:
          "Excel.FilterPivotHierarchyCollection.remove => (filterPivotHierarchy: Excel.FilterPivotHierarchy) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FormatProtection",
    apiList: [
      {
        name: "Excel.FormatProtection.formulaHidden",
        description:
          "Specifies if Excel hides the formula for the cells in the range. A `null` value indicates that the entire range doesn't have a uniform formula hidden setting.",
        kind: "Property",
        signature: "Excel.FormatProtection.formulaHidden: boolean",
        examples: [],
      },
      {
        name: "Excel.FormatProtection.locked",
        description:
          "Specifies if Excel locks the cells in the object. A `null` value indicates that the entire range doesn't have a uniform lock setting.",
        kind: "Property",
        signature: "Excel.FormatProtection.locked: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FormattedNumberCellValue",
    apiList: [
      {
        name: "Excel.FormattedNumberCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.FormattedNumberCellValue.basicType: RangeValueType.double | "Double"',
        examples: [],
      },
      {
        name: "Excel.FormattedNumberCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: "Excel.FormattedNumberCellValue.basicValue: number",
        examples: [],
      },
      {
        name: "Excel.FormattedNumberCellValue.numberFormat",
        description:
          "Returns the number format string that is used to display this value. When accessed through a `valuesAsJson` property, this number format string is in the en-US locale. When accessed through a `valuesAsJsonLocal` property, this number format is in the user's display locale. Number format strings must conform to Excel guidelines. To learn more, see Review guidelines for customizing a number format.",
        kind: "Property",
        signature: "Excel.FormattedNumberCellValue.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.FormattedNumberCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature:
          'Excel.FormattedNumberCellValue.type: CellValueType.formattedNumber | "FormattedNumber"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.FunctionResult",
    apiList: [
      {
        name: "Excel.FunctionResult.error",
        description:
          'Error value (such as "#DIV/0") representing the error. If the error string is not set, then the function succeeded, and its result is written to the Value field. The error is always in the English locale.',
        kind: "Property",
        signature: "Excel.FunctionResult.error: string",
        examples: [],
      },
      {
        name: "Excel.FunctionResult.value",
        description:
          "The value of function evaluation. The value field will be populated only if no error has occurred (i.e., the Error property is not set).",
        kind: "Property",
        signature: "Excel.FunctionResult.value: T",
        examples: [
          '" Number of wrenches sold in November = " + unitSoldInNov.value;',
          '" Number of wrenches sold in November and December = " + sumOfTwoLookups.value;',
        ],
      },
      {
        name: "Excel.FunctionResult",
        description: "An object containing the result of a function-evaluation operation",
        kind: "Class",
        signature:
          "Excel.FunctionResult<string | number | boolean>.value: string | number | boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Functions",
    apiList: [
      {
        name: "Excel.Functions.abs",
        description: "Returns the absolute value of a number, a number without its sign.",
        kind: "Method",
        signature:
          "Excel.Functions.abs => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.accrInt",
        description: "Returns the accrued interest for a security that pays periodic interest.",
        kind: "Method",
        signature:
          "Excel.Functions.accrInt => (issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstInterest: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: ...",
        examples: [],
      },
      {
        name: "Excel.Functions.accrIntM",
        description: "Returns the accrued interest for a security that pays interest at maturity.",
        kind: "Method",
        signature:
          "Excel.Functions.accrIntM => (issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, par: number | s...",
        examples: [],
      },
      {
        name: "Excel.Functions.acos",
        description:
          "Returns the arccosine of a number, in radians in the range 0 to Pi. The arccosine is the angle whose cosine is Number.",
        kind: "Method",
        signature:
          "Excel.Functions.acos => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.acosh",
        description: "Returns the inverse hyperbolic cosine of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.acosh => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.acot",
        description: "Returns the arccotangent of a number, in radians in the range 0 to Pi.",
        kind: "Method",
        signature:
          "Excel.Functions.acot => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.acoth",
        description: "Returns the inverse hyperbolic cotangent of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.acoth => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.amorDegrc",
        description:
          "Returns the prorated linear depreciation of an asset for each accounting period.",
        kind: "Method",
        signature:
          "Excel.Functions.amorDegrc => (cost: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, datePurchased: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvag...",
        examples: [],
      },
      {
        name: "Excel.Functions.amorLinc",
        description:
          "Returns the prorated linear depreciation of an asset for each accounting period.",
        kind: "Method",
        signature:
          "Excel.Functions.amorLinc => (cost: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, datePurchased: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvag...",
        examples: [],
      },
      {
        name: "Excel.Functions.and",
        description:
          "Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE.",
        kind: "Method",
        signature:
          "Excel.Functions.and => (...values: Array<boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.arabic",
        description: "Converts a Roman numeral to Arabic.",
        kind: "Method",
        signature:
          "Excel.Functions.arabic => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.areas",
        description:
          "Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell.",
        kind: "Method",
        signature:
          "Excel.Functions.areas => (reference: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.asc",
        description:
          "Changes full-width (double-byte) characters to half-width (single-byte) characters. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.asc => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.asin",
        description: "Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2.",
        kind: "Method",
        signature:
          "Excel.Functions.asin => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.asinh",
        description: "Returns the inverse hyperbolic sine of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.asinh => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.atan",
        description: "Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2.",
        kind: "Method",
        signature:
          "Excel.Functions.atan => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.atan2",
        description:
          "Returns the arctangent of the specified x- and y- coordinates, in radians between -Pi and Pi, excluding -Pi.",
        kind: "Method",
        signature:
          "Excel.Functions.atan2 => (xNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.atanh",
        description: "Returns the inverse hyperbolic tangent of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.atanh => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.aveDev",
        description:
          "Returns the average of the absolute deviations of data points from their mean. Arguments can be numbers or names, arrays, or references that contain numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.aveDev => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.average",
        description:
          "Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays, or references that contain numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.average => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.averageA",
        description:
          "Returns the average (arithmetic mean) of its arguments, evaluating text and FALSE in arguments as 0; TRUE evaluates as 1. Arguments can be numbers, names, arrays, or references.",
        kind: "Method",
        signature:
          "Excel.Functions.averageA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.averageIf",
        description:
          "Finds average(arithmetic mean) for the cells specified by a given condition or criteria.",
        kind: "Method",
        signature:
          "Excel.Functions.averageIf => (range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, averageRange?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.averageIfs",
        description:
          "Finds average(arithmetic mean) for the cells specified by a given set of conditions or criteria.",
        kind: "Method",
        signature:
          "Excel.Functions.averageIfs => (averageRange: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any> | number | string | boolean>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bahtText",
        description: "Converts a number to text (baht).",
        kind: "Method",
        signature:
          "Excel.Functions.bahtText => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.base",
        description: "Converts a number into a text representation with the given radix (base).",
        kind: "Method",
        signature:
          "Excel.Functions.base => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, radix: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, minLength?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.besselI",
        description: "Returns the modified Bessel function In(x).",
        kind: "Method",
        signature:
          "Excel.Functions.besselI => (x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.besselJ",
        description: "Returns the Bessel function Jn(x).",
        kind: "Method",
        signature:
          "Excel.Functions.besselJ => (x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.besselK",
        description: "Returns the modified Bessel function Kn(x).",
        kind: "Method",
        signature:
          "Excel.Functions.besselK => (x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.besselY",
        description: "Returns the Bessel function Yn(x).",
        kind: "Method",
        signature:
          "Excel.Functions.besselY => (x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.beta_Dist",
        description: "Returns the beta probability distribution function.",
        kind: "Method",
        signature:
          "Excel.Functions.beta_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<a...",
        examples: [],
      },
      {
        name: "Excel.Functions.beta_Inv",
        description:
          "Returns the inverse of the cumulative beta probability density function (BETA.DIST).",
        kind: "Method",
        signature:
          "Excel.Functions.beta_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, A?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<...",
        examples: [],
      },
      {
        name: "Excel.Functions.bin2Dec",
        description: "Converts a binary number to decimal.",
        kind: "Method",
        signature:
          "Excel.Functions.bin2Dec => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bin2Hex",
        description: "Converts a binary number to hexadecimal.",
        kind: "Method",
        signature:
          "Excel.Functions.bin2Hex => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bin2Oct",
        description: "Converts a binary number to octal.",
        kind: "Method",
        signature:
          "Excel.Functions.bin2Oct => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.binom_Dist",
        description: "Returns the individual term binomial distribution probability.",
        kind: "Method",
        signature:
          "Excel.Functions.binom_Dist => (numberS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, trials: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.F...",
        examples: [],
      },
      {
        name: "Excel.Functions.binom_Dist_Range",
        description: "Returns the probability of a trial result using a binomial distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.binom_Dist_Range => (trials: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberS2?: number | Excel.Range | Excel.RangeReference | Excel.Fun...",
        examples: [],
      },
      {
        name: "Excel.Functions.binom_Inv",
        description:
          "Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.",
        kind: "Method",
        signature:
          "Excel.Functions.binom_Inv => (trials: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bitand",
        description: "Returns a bitwise 'And' of two numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.bitand => (number1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bitlshift",
        description: "Returns a number shifted left by shift_amount bits.",
        kind: "Method",
        signature:
          "Excel.Functions.bitlshift => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, shiftAmount: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bitor",
        description: "Returns a bitwise 'Or' of two numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.bitor => (number1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bitrshift",
        description: "Returns a number shifted right by shift_amount bits.",
        kind: "Method",
        signature:
          "Excel.Functions.bitrshift => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, shiftAmount: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.bitxor",
        description: "Returns a bitwise 'Exclusive Or' of two numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.bitxor => (number1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ceiling_Math",
        description:
          "Rounds a number up, to the nearest integer or to the nearest multiple of significance.",
        kind: "Method",
        signature:
          "Excel.Functions.ceiling_Math => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mode?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ceiling_Precise",
        description:
          "Rounds a number up, to the nearest integer or to the nearest multiple of significance.",
        kind: "Method",
        signature:
          "Excel.Functions.ceiling_Precise => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.char",
        description:
          "Returns the character specified by the code number from the character set for your computer.",
        kind: "Method",
        signature:
          "Excel.Functions.char => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.chiSq_Dist",
        description: "Returns the left-tailed probability of the chi-squared distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.chiSq_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.chiSq_Dist_RT",
        description: "Returns the right-tailed probability of the chi-squared distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.chiSq_Dist_RT => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.chiSq_Inv",
        description:
          "Returns the inverse of the left-tailed probability of the chi-squared distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.chiSq_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.chiSq_Inv_RT",
        description:
          "Returns the inverse of the right-tailed probability of the chi-squared distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.chiSq_Inv_RT => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.choose",
        description:
          "Chooses a value or action to perform from a list of values, based on an index number.",
        kind: "Method",
        signature:
          "Excel.Functions.choose => (indexNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number | string | boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.clean",
        description: "Removes all nonprintable characters from text.",
        kind: "Method",
        signature:
          "Excel.Functions.clean => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.code",
        description:
          "Returns a numeric code for the first character in a text string, in the character set used by your computer.",
        kind: "Method",
        signature:
          "Excel.Functions.code => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.columns",
        description: "Returns the number of columns in an array or reference.",
        kind: "Method",
        signature:
          "Excel.Functions.columns => (array: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.combin",
        description: "Returns the number of combinations for a given number of items.",
        kind: "Method",
        signature:
          "Excel.Functions.combin => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.combina",
        description:
          "Returns the number of combinations with repetitions for a given number of items.",
        kind: "Method",
        signature:
          "Excel.Functions.combina => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.complex",
        description: "Converts real and imaginary coefficients into a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.complex => (realNum: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, iNum: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, suffix?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResul...",
        examples: [],
      },
      {
        name: "Excel.Functions.concatenate",
        description: "Joins several text strings into one text string.",
        kind: "Method",
        signature:
          "Excel.Functions.concatenate => (...values: Array<string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.confidence_Norm",
        description:
          "Returns the confidence interval for a population mean, using a normal distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.confidence_Norm => (alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, size: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.confidence_T",
        description:
          "Returns the confidence interval for a population mean, using a Student's T distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.confidence_T => (alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, size: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.convert",
        description: "Converts a number from one measurement system to another.",
        kind: "Method",
        signature:
          "Excel.Functions.convert => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fromUnit: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, toUnit: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionRes...",
        examples: [],
      },
      {
        name: "Excel.Functions.cos",
        description: "Returns the cosine of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.cos => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.cosh",
        description: "Returns the hyperbolic cosine of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.cosh => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.cot",
        description: "Returns the cotangent of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.cot => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.coth",
        description: "Returns the hyperbolic cotangent of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.coth => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.count",
        description: "Counts the number of cells in a range that contain numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.count => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.countA",
        description: "Counts the number of cells in a range that are not empty.",
        kind: "Method",
        signature:
          "Excel.Functions.countA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.countBlank",
        description: "Counts the number of empty cells in a specified range of cells.",
        kind: "Method",
        signature:
          "Excel.Functions.countBlank => (range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.countIf",
        description: "Counts the number of cells within a range that meet the given condition.",
        kind: "Method",
        signature:
          "Excel.Functions.countIf => (range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.countIfs",
        description:
          "Counts the number of cells specified by a given set of conditions or criteria.",
        kind: "Method",
        signature:
          "Excel.Functions.countIfs => (...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any> | number | string | boolean>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.coupDayBs",
        description:
          "Returns the number of days from the beginning of the coupon period to the settlement date.",
        kind: "Method",
        signature:
          "Excel.Functions.coupDayBs => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?:...",
        examples: [],
      },
      {
        name: "Excel.Functions.coupDays",
        description:
          "Returns the number of days in the coupon period that contains the settlement date.",
        kind: "Method",
        signature:
          "Excel.Functions.coupDays => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?:...",
        examples: [],
      },
      {
        name: "Excel.Functions.coupDaysNc",
        description: "Returns the number of days from the settlement date to the next coupon date.",
        kind: "Method",
        signature:
          "Excel.Functions.coupDaysNc => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?:...",
        examples: [],
      },
      {
        name: "Excel.Functions.coupNcd",
        description: "Returns the next coupon date after the settlement date.",
        kind: "Method",
        signature:
          "Excel.Functions.coupNcd => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?:...",
        examples: [],
      },
      {
        name: "Excel.Functions.coupNum",
        description:
          "Returns the number of coupons payable between the settlement date and maturity date.",
        kind: "Method",
        signature:
          "Excel.Functions.coupNum => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?:...",
        examples: [],
      },
      {
        name: "Excel.Functions.coupPcd",
        description: "Returns the previous coupon date before the settlement date.",
        kind: "Method",
        signature:
          "Excel.Functions.coupPcd => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?:...",
        examples: [],
      },
      {
        name: "Excel.Functions.csc",
        description: "Returns the cosecant of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.csc => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.csch",
        description: "Returns the hyperbolic cosecant of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.csch => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.cumIPmt",
        description: "Returns the cumulative interest paid between two periods.",
        kind: "Method",
        signature:
          "Excel.Functions.cumIPmt => (rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startPeriod: number | st...",
        examples: [],
      },
      {
        name: "Excel.Functions.cumPrinc",
        description: "Returns the cumulative principal paid on a loan between two periods.",
        kind: "Method",
        signature:
          "Excel.Functions.cumPrinc => (rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startPeriod: number | st...",
        examples: [],
      },
      {
        name: "Excel.Functions.date",
        description:
          "Returns the number that represents the date in Microsoft Excel date-time code.",
        kind: "Method",
        signature:
          "Excel.Functions.date => (year: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, month: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, day: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.datevalue",
        description:
          "Converts a date in the form of text to a number that represents the date in Microsoft Excel date-time code.",
        kind: "Method",
        signature:
          "Excel.Functions.datevalue => (dateText: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.daverage",
        description:
          "Averages the values in a column in a list or database that match conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.daverage => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.day",
        description: "Returns the day of the month, a number from 1 to 31.",
        kind: "Method",
        signature:
          "Excel.Functions.day => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.days",
        description: "Returns the number of days between the two dates.",
        kind: "Method",
        signature:
          "Excel.Functions.days => (endDate: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startDate: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.days360",
        description:
          "Returns the number of days between two dates based on a 360-day year (twelve 30-day months).",
        kind: "Method",
        signature:
          "Excel.Functions.days360 => (startDate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, method?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.db",
        description:
          "Returns the depreciation of an asset for a specified period using the fixed-declining balance method.",
        kind: "Method",
        signature:
          "Excel.Functions.db => (cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, period: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<a...",
        examples: [],
      },
      {
        name: "Excel.Functions.dbcs",
        description:
          "Changes half-width (single-byte) characters within a character string to full-width (double-byte) characters. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.dbcs => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.dcount",
        description:
          "Counts the cells containing numbers in the field (column) of records in the database that match the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dcount => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dcountA",
        description:
          "Counts nonblank cells in the field (column) of records in the database that match the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dcountA => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ddb",
        description:
          "Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.ddb => (cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, period: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<a...",
        examples: [],
      },
      {
        name: "Excel.Functions.dec2Bin",
        description: "Converts a decimal number to binary.",
        kind: "Method",
        signature:
          "Excel.Functions.dec2Bin => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dec2Hex",
        description: "Converts a decimal number to hexadecimal.",
        kind: "Method",
        signature:
          "Excel.Functions.dec2Hex => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dec2Oct",
        description: "Converts a decimal number to octal.",
        kind: "Method",
        signature:
          "Excel.Functions.dec2Oct => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.decimal",
        description:
          "Converts a text representation of a number in a given base into a decimal number.",
        kind: "Method",
        signature:
          "Excel.Functions.decimal => (number: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, radix: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.degrees",
        description: "Converts radians to degrees.",
        kind: "Method",
        signature:
          "Excel.Functions.degrees => (angle: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.delta",
        description: "Tests whether two numbers are equal.",
        kind: "Method",
        signature:
          "Excel.Functions.delta => (number1: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.devSq",
        description:
          "Returns the sum of squares of deviations of data points from their sample mean.",
        kind: "Method",
        signature:
          "Excel.Functions.devSq => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dget",
        description:
          "Extracts from a database a single record that matches the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dget => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number | boolean | string>",
        examples: [],
      },
      {
        name: "Excel.Functions.disc",
        description: "Returns the discount rate for a security.",
        kind: "Method",
        signature:
          "Excel.Functions.disc => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: nu...",
        examples: [],
      },
      {
        name: "Excel.Functions.dmax",
        description:
          "Returns the largest number in the field (column) of records in the database that match the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dmax => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dmin",
        description:
          "Returns the smallest number in the field (column) of records in the database that match the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dmin => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dollar",
        description: "Converts a number to text, using currency format.",
        kind: "Method",
        signature:
          "Excel.Functions.dollar => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimals?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.dollarDe",
        description:
          "Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number.",
        kind: "Method",
        signature:
          "Excel.Functions.dollarDe => (fractionalDollar: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fraction: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dollarFr",
        description:
          "Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction.",
        kind: "Method",
        signature:
          "Excel.Functions.dollarFr => (decimalDollar: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fraction: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dproduct",
        description:
          "Multiplies the values in the field (column) of records in the database that match the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dproduct => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dstDev",
        description:
          "Estimates the standard deviation based on a sample from selected database entries.",
        kind: "Method",
        signature:
          "Excel.Functions.dstDev => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dstDevP",
        description:
          "Calculates the standard deviation based on the entire population of selected database entries.",
        kind: "Method",
        signature:
          "Excel.Functions.dstDevP => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dsum",
        description:
          "Adds the numbers in the field (column) of records in the database that match the conditions you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.dsum => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.duration",
        description: "Returns the annual duration of a security with periodic interest payments.",
        kind: "Method",
        signature:
          "Excel.Functions.duration => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, coupon: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: numbe...",
        examples: [],
      },
      {
        name: "Excel.Functions.dvar",
        description: "Estimates variance based on a sample from selected database entries.",
        kind: "Method",
        signature:
          "Excel.Functions.dvar => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.dvarP",
        description:
          "Calculates variance based on the entire population of selected database entries.",
        kind: "Method",
        signature:
          "Excel.Functions.dvarP => (database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ecma_Ceiling",
        description:
          "Rounds a number up, to the nearest integer or to the nearest multiple of significance.",
        kind: "Method",
        signature:
          "Excel.Functions.ecma_Ceiling => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.edate",
        description:
          "Returns the serial number of the date that is the indicated number of months before or after the start date.",
        kind: "Method",
        signature:
          "Excel.Functions.edate => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, months: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.effect",
        description: "Returns the effective annual interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.effect => (nominalRate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, npery: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.eoMonth",
        description:
          "Returns the serial number of the last day of the month before or after a specified number of months.",
        kind: "Method",
        signature:
          "Excel.Functions.eoMonth => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, months: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.erf",
        description: "Returns the error function.",
        kind: "Method",
        signature:
          "Excel.Functions.erf => (lowerLimit: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, upperLimit?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.erf_Precise",
        description: "Returns the error function.",
        kind: "Method",
        signature:
          "Excel.Functions.erf_Precise => (X: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.erfC",
        description: "Returns the complementary error function.",
        kind: "Method",
        signature:
          "Excel.Functions.erfC => (x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.erfC_Precise",
        description: "Returns the complementary error function.",
        kind: "Method",
        signature:
          "Excel.Functions.erfC_Precise => (X: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.error_Type",
        description: "Returns a number matching an error value.",
        kind: "Method",
        signature:
          "Excel.Functions.error_Type => (errorVal: string | number | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.even",
        description:
          "Rounds a positive number up and negative number down to the nearest even integer.",
        kind: "Method",
        signature:
          "Excel.Functions.even => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.exact",
        description:
          "Checks whether two text strings are exactly the same, and returns TRUE or FALSE. EXACT is case-sensitive.",
        kind: "Method",
        signature:
          "Excel.Functions.exact => (text1: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, text2: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.exp",
        description: "Returns e raised to the power of a given number.",
        kind: "Method",
        signature:
          "Excel.Functions.exp => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.expon_Dist",
        description: "Returns the exponential distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.expon_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lambda: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.f_Dist",
        description:
          "Returns the (left-tailed) F probability distribution (degree of diversity) for two data sets.",
        kind: "Method",
        signature:
          "Excel.Functions.f_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.Fun...",
        examples: [],
      },
      {
        name: "Excel.Functions.f_Dist_RT",
        description:
          "Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets.",
        kind: "Method",
        signature:
          "Excel.Functions.f_Dist_RT => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.f_Inv",
        description:
          "Returns the inverse of the (left-tailed) F probability distribution: if p = F.DIST(x,...), then F.INV(p,...) = x.",
        kind: "Method",
        signature:
          "Excel.Functions.f_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.f_Inv_RT",
        description:
          "Returns the inverse of the (right-tailed) F probability distribution: if p = F.DIST.RT(x,...), then F.INV.RT(p,...) = x.",
        kind: "Method",
        signature:
          "Excel.Functions.f_Inv_RT => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.fact",
        description: "Returns the factorial of a number, equal to 1*2*3*...* Number.",
        kind: "Method",
        signature:
          "Excel.Functions.fact => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.factDouble",
        description: "Returns the double factorial of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.factDouble => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.false",
        description: "Returns the logical value FALSE.",
        kind: "Method",
        signature: "Excel.Functions.false => () => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.find",
        description:
          "Returns the starting position of one text string within another text string. FIND is case-sensitive.",
        kind: "Method",
        signature:
          "Excel.Functions.find => (findText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, withinText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.findB",
        description:
          "Finds the starting position of one text string within another text string. FINDB is case-sensitive. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.findB => (findText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, withinText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.fisher",
        description: "Returns the Fisher transformation.",
        kind: "Method",
        signature:
          "Excel.Functions.fisher => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.fisherInv",
        description:
          "Returns the inverse of the Fisher transformation: if y = FISHER(x), then FISHERINV(y) = x.",
        kind: "Method",
        signature:
          "Excel.Functions.fisherInv => (y: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.fixed",
        description:
          "Rounds a number to the specified number of decimals and returns the result as text with or without commas.",
        kind: "Method",
        signature:
          "Excel.Functions.fixed => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimals?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, noCommas?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.floor_Math",
        description:
          "Rounds a number down, to the nearest integer or to the nearest multiple of significance.",
        kind: "Method",
        signature:
          "Excel.Functions.floor_Math => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mode?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.floor_Precise",
        description:
          "Rounds a number down, to the nearest integer or to the nearest multiple of significance.",
        kind: "Method",
        signature:
          "Excel.Functions.floor_Precise => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.fv",
        description:
          "Returns the future value of an investment based on periodic, constant payments and a constant interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.fv => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ty...",
        examples: [],
      },
      {
        name: "Excel.Functions.fvschedule",
        description:
          "Returns the future value of an initial principal after applying a series of compound interest rates.",
        kind: "Method",
        signature:
          "Excel.Functions.fvschedule => (principal: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, schedule: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.gamma",
        description: "Returns the Gamma function value.",
        kind: "Method",
        signature:
          "Excel.Functions.gamma => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.gamma_Dist",
        description: "Returns the gamma distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.gamma_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<a...",
        examples: [],
      },
      {
        name: "Excel.Functions.gamma_Inv",
        description:
          "Returns the inverse of the gamma cumulative distribution: if p = GAMMA.DIST(x,...), then GAMMA.INV(p,...) = x.",
        kind: "Method",
        signature:
          "Excel.Functions.gamma_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.gammaLn",
        description: "Returns the natural logarithm of the gamma function.",
        kind: "Method",
        signature:
          "Excel.Functions.gammaLn => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.gammaLn_Precise",
        description: "Returns the natural logarithm of the gamma function.",
        kind: "Method",
        signature:
          "Excel.Functions.gammaLn_Precise => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.gauss",
        description: "Returns 0.5 less than the standard normal cumulative distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.gauss => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.gcd",
        description: "Returns the greatest common divisor.",
        kind: "Method",
        signature:
          "Excel.Functions.gcd => (...values: Array<number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.geoMean",
        description: "Returns the geometric mean of an array or range of positive numeric data.",
        kind: "Method",
        signature:
          "Excel.Functions.geoMean => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.geStep",
        description: "Tests whether a number is greater than a threshold value.",
        kind: "Method",
        signature:
          "Excel.Functions.geStep => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, step?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.harMean",
        description:
          "Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals.",
        kind: "Method",
        signature:
          "Excel.Functions.harMean => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.hex2Bin",
        description: "Converts a Hexadecimal number to binary.",
        kind: "Method",
        signature:
          "Excel.Functions.hex2Bin => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.hex2Dec",
        description: "Converts a hexadecimal number to decimal.",
        kind: "Method",
        signature:
          "Excel.Functions.hex2Dec => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.hex2Oct",
        description: "Converts a hexadecimal number to octal.",
        kind: "Method",
        signature:
          "Excel.Functions.hex2Oct => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.hlookup",
        description:
          "Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.hlookup => (lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, tableArray: Excel.Range | number | Excel.RangeReference | Excel.FunctionResult<any>, rowIndexNum: Excel.Range | number | Excel.RangeReference | Excel.FunctionResult<any>, rangeLookup?: boolean | Excel.Range | Ex...",
        examples: [],
      },
      {
        name: "Excel.Functions.hour",
        description: "Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.).",
        kind: "Method",
        signature:
          "Excel.Functions.hour => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.hyperlink",
        description:
          "Creates a shortcut or jump that opens a document stored on your hard drive, a network server, or on the Internet.",
        kind: "Method",
        signature:
          "Excel.Functions.hyperlink => (linkLocation: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, friendlyName?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number | string | boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.hypGeom_Dist",
        description: "Returns the hypergeometric distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.hypGeom_Dist => (sampleS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberSample: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, populationS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberPop: number | Excel.Range | Excel.RangeReference | Exce...",
        examples: [],
      },
      {
        name: "Excel.Functions.if",
        description:
          "Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.if => (logicalTest: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, valueIfTrue?: Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>, valueIfFalse?: Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResul...",
        examples: [],
      },
      {
        name: "Excel.Functions.imAbs",
        description: "Returns the absolute value (modulus) of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imAbs => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imaginary",
        description: "Returns the imaginary coefficient of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imaginary => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imArgument",
        description: "Returns the argument q, an angle expressed in radians.",
        kind: "Method",
        signature:
          "Excel.Functions.imArgument => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imConjugate",
        description: "Returns the complex conjugate of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imConjugate => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imCos",
        description: "Returns the cosine of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imCos => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imCosh",
        description: "Returns the hyperbolic cosine of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imCosh => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imCot",
        description: "Returns the cotangent of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imCot => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imCsc",
        description: "Returns the cosecant of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imCsc => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imCsch",
        description: "Returns the hyperbolic cosecant of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imCsch => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imDiv",
        description: "Returns the quotient of two complex numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.imDiv => (inumber1: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, inumber2: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imExp",
        description: "Returns the exponential of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imExp => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imLn",
        description: "Returns the natural logarithm of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imLn => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imLog10",
        description: "Returns the base-10 logarithm of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imLog10 => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imLog2",
        description: "Returns the base-2 logarithm of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imLog2 => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imPower",
        description: "Returns a complex number raised to an integer power.",
        kind: "Method",
        signature:
          "Excel.Functions.imPower => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imProduct",
        description: "Returns the product of 1 to 255 complex numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.imProduct => (...values: Array<Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imReal",
        description: "Returns the real coefficient of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imReal => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSec",
        description: "Returns the secant of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imSec => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSech",
        description: "Returns the hyperbolic secant of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imSech => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSin",
        description: "Returns the sine of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imSin => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSinh",
        description: "Returns the hyperbolic sine of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imSinh => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSqrt",
        description: "Returns the square root of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imSqrt => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSub",
        description: "Returns the difference of two complex numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.imSub => (inumber1: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, inumber2: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imSum",
        description: "Returns the sum of complex numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.imSum => (...values: Array<Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.imTan",
        description: "Returns the tangent of a complex number.",
        kind: "Method",
        signature:
          "Excel.Functions.imTan => (inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.int",
        description: "Rounds a number down to the nearest integer.",
        kind: "Method",
        signature:
          "Excel.Functions.int => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.intRate",
        description: "Returns the interest rate for a fully invested security.",
        kind: "Method",
        signature:
          "Excel.Functions.intRate => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, investment: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemp...",
        examples: [],
      },
      {
        name: "Excel.Functions.ipmt",
        description:
          "Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.ipmt => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?...",
        examples: [],
      },
      {
        name: "Excel.Functions.irr",
        description: "Returns the internal rate of return for a series of cash flows.",
        kind: "Method",
        signature:
          "Excel.Functions.irr => (values: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, guess?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.isErr",
        description:
          "Checks whether a value is an error (#VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!) excluding #N/A, and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isErr => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isError",
        description:
          "Checks whether a value is an error (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!), and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isError => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isEven",
        description: "Returns TRUE if the number is even.",
        kind: "Method",
        signature:
          "Excel.Functions.isEven => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.isFormula",
        description:
          "Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isFormula => (reference: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isLogical",
        description:
          "Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isLogical => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isNA",
        description: "Checks whether a value is #N/A, and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isNA => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isNonText",
        description:
          "Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isNonText => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isNumber",
        description: "Checks whether a value is a number, and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isNumber => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.iso_Ceiling",
        description:
          "Rounds a number up, to the nearest integer or to the nearest multiple of significance.",
        kind: "Method",
        signature:
          "Excel.Functions.iso_Ceiling => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.isOdd",
        description: "Returns TRUE if the number is odd.",
        kind: "Method",
        signature:
          "Excel.Functions.isOdd => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.isoWeekNum",
        description: "Returns the ISO week number in the year for a given date.",
        kind: "Method",
        signature:
          "Excel.Functions.isoWeekNum => (date: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ispmt",
        description: "Returns the interest paid during a specific period of an investment.",
        kind: "Method",
        signature:
          "Excel.Functions.ispmt => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => ...",
        examples: [],
      },
      {
        name: "Excel.Functions.isref",
        description: "Checks whether a value is a reference, and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isref => (value: Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.isText",
        description: "Checks whether a value is text, and returns TRUE or FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.isText => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.kurt",
        description: "Returns the kurtosis of a data set.",
        kind: "Method",
        signature:
          "Excel.Functions.kurt => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.large",
        description:
          "Returns the k-th largest value in a data set. For example, the fifth largest number.",
        kind: "Method",
        signature:
          "Excel.Functions.large => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.lcm",
        description: "Returns the least common multiple.",
        kind: "Method",
        signature:
          "Excel.Functions.lcm => (...values: Array<number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.left",
        description: "Returns the specified number of characters from the start of a text string.",
        kind: "Method",
        signature:
          "Excel.Functions.left => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.leftb",
        description:
          "Returns the specified number of characters from the start of a text string. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.leftb => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.len",
        description: "Returns the number of characters in a text string.",
        kind: "Method",
        signature:
          "Excel.Functions.len => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.lenb",
        description:
          "Returns the number of characters in a text string. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.lenb => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ln",
        description: "Returns the natural logarithm of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.ln => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.log",
        description: "Returns the logarithm of a number to the base you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.log => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, base?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.log10",
        description: "Returns the base-10 logarithm of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.log10 => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.logNorm_Dist",
        description:
          "Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.",
        kind: "Method",
        signature:
          "Excel.Functions.logNorm_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionRe...",
        examples: [],
      },
      {
        name: "Excel.Functions.logNorm_Inv",
        description:
          "Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.",
        kind: "Method",
        signature:
          "Excel.Functions.logNorm_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.lookup",
        description:
          "Looks up a value either from a one-row or one-column range or from an array. Provided for backward compatibility.",
        kind: "Method",
        signature:
          "Excel.Functions.lookup => (lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lookupVector: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, resultVector?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number | string | boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.lower",
        description: "Converts all letters in a text string to lowercase.",
        kind: "Method",
        signature:
          "Excel.Functions.lower => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.match",
        description:
          "Returns the relative position of an item in an array that matches a specified value in a specified order.",
        kind: "Method",
        signature:
          "Excel.Functions.match => (lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lookupArray: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, matchType?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.max",
        description:
          "Returns the largest value in a set of values. Ignores logical values and text.",
        kind: "Method",
        signature:
          "Excel.Functions.max => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.maxA",
        description:
          "Returns the largest value in a set of values. Does not ignore logical values and text.",
        kind: "Method",
        signature:
          "Excel.Functions.maxA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.mduration",
        description:
          "Returns the Macauley modified duration for a security with an assumed par value of $100.",
        kind: "Method",
        signature:
          "Excel.Functions.mduration => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, coupon: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: numbe...",
        examples: [],
      },
      {
        name: "Excel.Functions.median",
        description: "Returns the median, or the number in the middle of the set of given numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.median => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.mid",
        description:
          "Returns the characters from the middle of a text string, given a starting position and length.",
        kind: "Method",
        signature:
          "Excel.Functions.mid => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.midb",
        description:
          "Returns characters from the middle of a text string, given a starting position and length. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.midb => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.min",
        description:
          "Returns the smallest number in a set of values. Ignores logical values and text.",
        kind: "Method",
        signature:
          "Excel.Functions.min => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.minA",
        description:
          "Returns the smallest value in a set of values. Does not ignore logical values and text.",
        kind: "Method",
        signature:
          "Excel.Functions.minA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.minute",
        description: "Returns the minute, a number from 0 to 59.",
        kind: "Method",
        signature:
          "Excel.Functions.minute => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.mirr",
        description:
          "Returns the internal rate of return for a series of periodic cash flows, considering both cost of investment and interest on reinvestment of cash.",
        kind: "Method",
        signature:
          "Excel.Functions.mirr => (values: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, financeRate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, reinvestRate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.mod",
        description: "Returns the remainder after a number is divided by a divisor.",
        kind: "Method",
        signature:
          "Excel.Functions.mod => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, divisor: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.month",
        description: "Returns the month, a number from 1 (January) to 12 (December).",
        kind: "Method",
        signature:
          "Excel.Functions.month => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.mround",
        description: "Returns a number rounded to the desired multiple.",
        kind: "Method",
        signature:
          "Excel.Functions.mround => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, multiple: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.multiNomial",
        description: "Returns the multinomial of a set of numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.multiNomial => (...values: Array<number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.n",
        description:
          "Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero).",
        kind: "Method",
        signature:
          "Excel.Functions.n => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.na",
        description: "Returns the error value #N/A (value not available).",
        kind: "Method",
        signature: "Excel.Functions.na => () => FunctionResult<number | string>",
        examples: [],
      },
      {
        name: "Excel.Functions.negBinom_Dist",
        description:
          "Returns the negative binomial distribution, the probability that there will be Number_f failures before the Number_s-th success, with Probability_s probability of a success.",
        kind: "Method",
        signature:
          "Excel.Functions.negBinom_Dist => (numberF: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel....",
        examples: [],
      },
      {
        name: "Excel.Functions.networkDays",
        description: "Returns the number of whole workdays between two dates.",
        kind: "Method",
        signature:
          "Excel.Functions.networkDays => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>) => Functi...",
        examples: [],
      },
      {
        name: "Excel.Functions.networkDays_Intl",
        description:
          "Returns the number of whole workdays between two dates with custom weekend parameters.",
        kind: "Method",
        signature:
          "Excel.Functions.networkDays_Intl => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, weekend?: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | ...",
        examples: [],
      },
      {
        name: "Excel.Functions.nominal",
        description: "Returns the annual nominal interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.nominal => (effectRate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, npery: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.norm_Dist",
        description:
          "Returns the normal distribution for the specified mean and standard deviation.",
        kind: "Method",
        signature:
          "Excel.Functions.norm_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionRe...",
        examples: [],
      },
      {
        name: "Excel.Functions.norm_Inv",
        description:
          "Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.",
        kind: "Method",
        signature:
          "Excel.Functions.norm_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.norm_S_Dist",
        description:
          "Returns the standard normal distribution (has a mean of zero and a standard deviation of one).",
        kind: "Method",
        signature:
          "Excel.Functions.norm_S_Dist => (z: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.norm_S_Inv",
        description:
          "Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).",
        kind: "Method",
        signature:
          "Excel.Functions.norm_S_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.not",
        description: "Changes FALSE to TRUE, or TRUE to FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.not => (logical: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.now",
        description: "Returns the current date and time formatted as a date and time.",
        kind: "Method",
        signature: "Excel.Functions.now => () => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.nper",
        description:
          "Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.nper => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type...",
        examples: [],
      },
      {
        name: "Excel.Functions.npv",
        description:
          "Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).",
        kind: "Method",
        signature:
          "Excel.Functions.npv => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.numberValue",
        description: "Converts text to number in a locale-independent manner.",
        kind: "Method",
        signature:
          "Excel.Functions.numberValue => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimalSeparator?: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, groupSeparator?: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.oct2Bin",
        description: "Converts an octal number to binary.",
        kind: "Method",
        signature:
          "Excel.Functions.oct2Bin => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.oct2Dec",
        description: "Converts an octal number to decimal.",
        kind: "Method",
        signature:
          "Excel.Functions.oct2Dec => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.oct2Hex",
        description: "Converts an octal number to hexadecimal.",
        kind: "Method",
        signature:
          "Excel.Functions.oct2Hex => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.odd",
        description:
          "Rounds a positive number up and negative number down to the nearest odd integer.",
        kind: "Method",
        signature:
          "Excel.Functions.odd => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.oddFPrice",
        description:
          "Returns the price per $100 face value of a security with an odd first period.",
        kind: "Method",
        signature:
          "Excel.Functions.oddFPrice => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstCoupon...",
        examples: [],
      },
      {
        name: "Excel.Functions.oddFYield",
        description: "Returns the yield of a security with an odd first period.",
        kind: "Method",
        signature:
          "Excel.Functions.oddFYield => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstCoupon...",
        examples: [],
      },
      {
        name: "Excel.Functions.oddLPrice",
        description: "Returns the price per $100 face value of a security with an odd last period.",
        kind: "Method",
        signature:
          "Excel.Functions.oddLPrice => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lastInterest: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate...",
        examples: [],
      },
      {
        name: "Excel.Functions.oddLYield",
        description: "Returns the yield of a security with an odd last period.",
        kind: "Method",
        signature:
          "Excel.Functions.oddLYield => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lastInterest: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate...",
        examples: [],
      },
      {
        name: "Excel.Functions.or",
        description:
          "Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE.",
        kind: "Method",
        signature:
          "Excel.Functions.or => (...values: Array<boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.pduration",
        description:
          "Returns the number of periods required by an investment to reach a specified value.",
        kind: "Method",
        signature:
          "Excel.Functions.pduration => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.percentile_Exc",
        description:
          "Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.",
        kind: "Method",
        signature:
          "Excel.Functions.percentile_Exc => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.percentile_Inc",
        description:
          "Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive.",
        kind: "Method",
        signature:
          "Excel.Functions.percentile_Inc => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.percentRank_Exc",
        description:
          "Returns the rank of a value in a data set as a percentage of the data set as a percentage (0..1, exclusive) of the data set.",
        kind: "Method",
        signature:
          "Excel.Functions.percentRank_Exc => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.percentRank_Inc",
        description:
          "Returns the rank of a value in a data set as a percentage of the data set as a percentage (0..1, inclusive) of the data set.",
        kind: "Method",
        signature:
          "Excel.Functions.percentRank_Inc => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.permut",
        description:
          "Returns the number of permutations for a given number of objects that can be selected from the total objects.",
        kind: "Method",
        signature:
          "Excel.Functions.permut => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.permutationa",
        description:
          "Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.",
        kind: "Method",
        signature:
          "Excel.Functions.permutationa => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.phi",
        description:
          "Returns the value of the density function for a standard normal distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.phi => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.pi",
        description: "Returns the value of Pi, 3.14159265358979, accurate to 15 digits.",
        kind: "Method",
        signature: "Excel.Functions.pi => () => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.pmt",
        description:
          "Calculates the payment for a loan based on constant payments and a constant interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.pmt => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, typ...",
        examples: [],
      },
      {
        name: "Excel.Functions.poisson_Dist",
        description: "Returns the Poisson distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.poisson_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.power",
        description: "Returns the result of a number raised to a power.",
        kind: "Method",
        signature:
          "Excel.Functions.power => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, power: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.ppmt",
        description:
          "Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.",
        kind: "Method",
        signature:
          "Excel.Functions.ppmt => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?...",
        examples: [],
      },
      {
        name: "Excel.Functions.price",
        description:
          "Returns the price per $100 face value of a security that pays periodic interest.",
        kind: "Method",
        signature:
          "Excel.Functions.price => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number ...",
        examples: [],
      },
      {
        name: "Excel.Functions.priceDisc",
        description: "Returns the price per $100 face value of a discounted security.",
        kind: "Method",
        signature:
          "Excel.Functions.priceDisc => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redempti...",
        examples: [],
      },
      {
        name: "Excel.Functions.priceMat",
        description:
          "Returns the price per $100 face value of a security that pays interest at maturity.",
        kind: "Method",
        signature:
          "Excel.Functions.priceMat => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: numbe...",
        examples: [],
      },
      {
        name: "Excel.Functions.product",
        description: "Multiplies all the numbers given as arguments.",
        kind: "Method",
        signature:
          "Excel.Functions.product => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.proper",
        description:
          "Converts a text string to proper case; the first letter in each word to uppercase, and all other letters to lowercase.",
        kind: "Method",
        signature:
          "Excel.Functions.proper => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.pv",
        description:
          "Returns the present value of an investment: the total amount that a series of future payments is worth now.",
        kind: "Method",
        signature:
          "Excel.Functions.pv => (rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ty...",
        examples: [],
      },
      {
        name: "Excel.Functions.quartile_Exc",
        description:
          "Returns the quartile of a data set, based on percentile values from 0..1, exclusive.",
        kind: "Method",
        signature:
          "Excel.Functions.quartile_Exc => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, quart: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.quartile_Inc",
        description:
          "Returns the quartile of a data set, based on percentile values from 0..1, inclusive.",
        kind: "Method",
        signature:
          "Excel.Functions.quartile_Inc => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, quart: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.quotient",
        description: "Returns the integer portion of a division.",
        kind: "Method",
        signature:
          "Excel.Functions.quotient => (numerator: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, denominator: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.radians",
        description: "Converts degrees to radians.",
        kind: "Method",
        signature:
          "Excel.Functions.radians => (angle: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.rand",
        description:
          "Returns a random number greater than or equal to 0 and less than 1, evenly distributed (changes on recalculation).",
        kind: "Method",
        signature: "Excel.Functions.rand => () => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.randBetween",
        description: "Returns a random number between the numbers you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.randBetween => (bottom: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, top: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.rank_Avg",
        description:
          "Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the average rank is returned.",
        kind: "Method",
        signature:
          "Excel.Functions.rank_Avg => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ref: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, order?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.rank_Eq",
        description:
          "Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the top rank of that set of values is returned.",
        kind: "Method",
        signature:
          "Excel.Functions.rank_Eq => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ref: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, order?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.rate",
        description:
          "Returns the interest rate per period of a loan or an investment. For example, use 6%/4 for quarterly payments at 6% APR.",
        kind: "Method",
        signature:
          "Excel.Functions.rate => (nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type...",
        examples: [],
      },
      {
        name: "Excel.Functions.received",
        description: "Returns the amount received at maturity for a fully invested security.",
        kind: "Method",
        signature:
          "Excel.Functions.received => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, investment: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discou...",
        examples: [],
      },
      {
        name: "Excel.Functions.replace",
        description: "Replaces part of a text string with a different text string.",
        kind: "Method",
        signature:
          "Excel.Functions.replace => (oldText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, newText: string | Excel.Range | Excel.RangeReference | Excel.Functio...",
        examples: [],
      },
      {
        name: "Excel.Functions.replaceB",
        description:
          "Replaces part of a text string with a different text string. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.replaceB => (oldText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, newText: string | Excel.Range | Excel.RangeReference | Excel.Functio...",
        examples: [],
      },
      {
        name: "Excel.Functions.rept",
        description:
          "Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string.",
        kind: "Method",
        signature:
          "Excel.Functions.rept => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberTimes: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.right",
        description: "Returns the specified number of characters from the end of a text string.",
        kind: "Method",
        signature:
          "Excel.Functions.right => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.rightb",
        description:
          "Returns the specified number of characters from the end of a text string. Use with double-byte character sets (DBCS).",
        kind: "Method",
        signature:
          "Excel.Functions.rightb => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.roman",
        description: "Converts an Arabic numeral to Roman, as text.",
        kind: "Method",
        signature:
          "Excel.Functions.roman => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, form?: boolean | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.round",
        description: "Rounds a number to a specified number of digits.",
        kind: "Method",
        signature:
          "Excel.Functions.round => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.roundDown",
        description: "Rounds a number down, toward zero.",
        kind: "Method",
        signature:
          "Excel.Functions.roundDown => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.roundUp",
        description: "Rounds a number up, away from zero.",
        kind: "Method",
        signature:
          "Excel.Functions.roundUp => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.rows",
        description: "Returns the number of rows in a reference or array.",
        kind: "Method",
        signature:
          "Excel.Functions.rows => (array: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.rri",
        description: "Returns an equivalent interest rate for the growth of an investment.",
        kind: "Method",
        signature:
          "Excel.Functions.rri => (nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sec",
        description: "Returns the secant of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.sec => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sech",
        description: "Returns the hyperbolic secant of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.sech => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.second",
        description: "Returns the second, a number from 0 to 59.",
        kind: "Method",
        signature:
          "Excel.Functions.second => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.seriesSum",
        description: "Returns the sum of a power series based on the formula.",
        kind: "Method",
        signature:
          "Excel.Functions.seriesSum => (x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, m: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, coefficients: Excel.Range | str...",
        examples: [],
      },
      {
        name: "Excel.Functions.sheet",
        description: "Returns the sheet number of the referenced sheet.",
        kind: "Method",
        signature:
          "Excel.Functions.sheet => (value?: Excel.Range | string | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sheets",
        description: "Returns the number of sheets in a reference.",
        kind: "Method",
        signature:
          "Excel.Functions.sheets => (reference?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sign",
        description:
          "Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative.",
        kind: "Method",
        signature:
          "Excel.Functions.sign => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sin",
        description: "Returns the sine of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.sin => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sinh",
        description: "Returns the hyperbolic sine of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.sinh => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.skew",
        description:
          "Returns the skewness of a distribution: a characterization of the degree of asymmetry of a distribution around its mean.",
        kind: "Method",
        signature:
          "Excel.Functions.skew => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.skew_p",
        description:
          "Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean.",
        kind: "Method",
        signature:
          "Excel.Functions.skew_p => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sln",
        description: "Returns the straight-line depreciation of an asset for one period.",
        kind: "Method",
        signature:
          "Excel.Functions.sln => (cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.small",
        description:
          "Returns the k-th smallest value in a data set. For example, the fifth smallest number.",
        kind: "Method",
        signature:
          "Excel.Functions.small => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sqrt",
        description: "Returns the square root of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.sqrt => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sqrtPi",
        description: "Returns the square root of (number * Pi).",
        kind: "Method",
        signature:
          "Excel.Functions.sqrtPi => (number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.standardize",
        description:
          "Returns a normalized value from a distribution characterized by a mean and standard deviation.",
        kind: "Method",
        signature:
          "Excel.Functions.standardize => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.stDev_P",
        description:
          "Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).",
        kind: "Method",
        signature:
          "Excel.Functions.stDev_P => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.stDev_S",
        description:
          "Estimates standard deviation based on a sample (ignores logical values and text in the sample).",
        kind: "Method",
        signature:
          "Excel.Functions.stDev_S => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.stDevA",
        description:
          "Estimates standard deviation based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.",
        kind: "Method",
        signature:
          "Excel.Functions.stDevA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.stDevPA",
        description:
          "Calculates standard deviation based on an entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.",
        kind: "Method",
        signature:
          "Excel.Functions.stDevPA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.substitute",
        description: "Replaces existing text with new text in a text string.",
        kind: "Method",
        signature:
          "Excel.Functions.substitute => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, oldText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, newText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, instanceNum?: string | Excel.Range | Excel.RangeReference | Excel.Functio...",
        examples: [],
      },
      {
        name: "Excel.Functions.subtotal",
        description: "Returns a subtotal in a list or database.",
        kind: "Method",
        signature:
          "Excel.Functions.subtotal => (functionNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sum",
        description: "Adds all the numbers in a range of cells.",
        kind: "Method",
        signature:
          "Excel.Functions.sum(...values: (number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>)[]) => Excel.FunctionResult<number>",
        examples: [
          'let sumOfTwoLookups = workbook.functions.sum(\n    workbook.functions.vlookup("Wrench", range, 2, false),\n    workbook.functions.vlookup("Wrench", range, 3, false)\n  );',
        ],
      },
      {
        name: "Excel.Functions.sumIf",
        description: "Adds the cells specified by a given condition or criteria.",
        kind: "Method",
        signature:
          "Excel.Functions.sumIf => (range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, sumRange?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sumIfs",
        description: "Adds the cells specified by a given set of conditions or criteria.",
        kind: "Method",
        signature:
          "Excel.Functions.sumIfs => (sumRange: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any> | number | string | boolean>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.sumSq",
        description:
          "Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names, or references to cells that contain numbers.",
        kind: "Method",
        signature:
          "Excel.Functions.sumSq => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.syd",
        description:
          "Returns the sum-of-years' digits depreciation of an asset for a specified period.",
        kind: "Method",
        signature:
          "Excel.Functions.syd => (cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>...",
        examples: [],
      },
      {
        name: "Excel.Functions.t",
        description:
          "Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not.",
        kind: "Method",
        signature:
          "Excel.Functions.t => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.t_Dist",
        description: "Returns the left-tailed Student's t-distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.t_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.t_Dist_2T",
        description: "Returns the two-tailed Student's t-distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.t_Dist_2T => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.t_Dist_RT",
        description: "Returns the right-tailed Student's t-distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.t_Dist_RT => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.t_Inv",
        description: "Returns the left-tailed inverse of the Student's t-distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.t_Inv => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.t_Inv_2T",
        description: "Returns the two-tailed inverse of the Student's t-distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.t_Inv_2T => (probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.tan",
        description: "Returns the tangent of an angle.",
        kind: "Method",
        signature:
          "Excel.Functions.tan => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.tanh",
        description: "Returns the hyperbolic tangent of a number.",
        kind: "Method",
        signature:
          "Excel.Functions.tanh => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.tbillEq",
        description: "Returns the bond-equivalent yield for a treasury bill.",
        kind: "Method",
        signature:
          "Excel.Functions.tbillEq => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => Funct...",
        examples: [],
      },
      {
        name: "Excel.Functions.tbillPrice",
        description: "Returns the price per $100 face value for a treasury bill.",
        kind: "Method",
        signature:
          "Excel.Functions.tbillPrice => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => Funct...",
        examples: [],
      },
      {
        name: "Excel.Functions.tbillYield",
        description: "Returns the yield for a treasury bill.",
        kind: "Method",
        signature:
          "Excel.Functions.tbillYield => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionRes...",
        examples: [],
      },
      {
        name: "Excel.Functions.text",
        description: "Converts a value to text in a specific number format.",
        kind: "Method",
        signature:
          "Excel.Functions.text => (value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, formatText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.time",
        description:
          "Converts hours, minutes, and seconds given as numbers to an Excel serial number, formatted with a time format.",
        kind: "Method",
        signature:
          "Excel.Functions.time => (hour: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, minute: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, second: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.timevalue",
        description:
          "Converts a text time to an Excel serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM). Format the number with a time format after entering the formula.",
        kind: "Method",
        signature:
          "Excel.Functions.timevalue => (timeText: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.today",
        description: "Returns the current date formatted as a date.",
        kind: "Method",
        signature: "Excel.Functions.today => () => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.trim",
        description:
          "Removes all spaces from a text string except for single spaces between words.",
        kind: "Method",
        signature:
          "Excel.Functions.trim => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.trimMean",
        description: "Returns the mean of the interior portion of a set of data values.",
        kind: "Method",
        signature:
          "Excel.Functions.trimMean => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, percent: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.true",
        description: "Returns the logical value TRUE.",
        kind: "Method",
        signature: "Excel.Functions.true => () => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.trunc",
        description:
          "Truncates a number to an integer by removing the decimal, or fractional, part of the number.",
        kind: "Method",
        signature:
          "Excel.Functions.trunc => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.type",
        description:
          "Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64.",
        kind: "Method",
        signature:
          "Excel.Functions.type => (value: boolean | string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.unichar",
        description: "Returns the Unicode character referenced by the given numeric value.",
        kind: "Method",
        signature:
          "Excel.Functions.unichar => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.unicode",
        description:
          "Returns the number (code point) corresponding to the first character of the text.",
        kind: "Method",
        signature:
          "Excel.Functions.unicode => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.upper",
        description: "Converts a text string to all uppercase letters.",
        kind: "Method",
        signature:
          "Excel.Functions.upper => (text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.usdollar",
        description: "Converts a number to text, using currency format.",
        kind: "Method",
        signature:
          "Excel.Functions.usdollar => (number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimals?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<string>",
        examples: [],
      },
      {
        name: "Excel.Functions.value",
        description: "Converts a text string that represents a number to a number.",
        kind: "Method",
        signature:
          "Excel.Functions.value => (text: string | boolean | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.var_P",
        description:
          "Calculates variance based on the entire population (ignores logical values and text in the population).",
        kind: "Method",
        signature:
          "Excel.Functions.var_P => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.var_S",
        description:
          "Estimates variance based on a sample (ignores logical values and text in the sample).",
        kind: "Method",
        signature:
          "Excel.Functions.var_S => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.varA",
        description:
          "Estimates variance based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.",
        kind: "Method",
        signature:
          "Excel.Functions.varA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.varPA",
        description:
          "Calculates variance based on the entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.",
        kind: "Method",
        signature:
          "Excel.Functions.varPA => (...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.vdb",
        description:
          "Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify.",
        kind: "Method",
        signature:
          "Excel.Functions.vdb => (cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startPeriod: number | Excel.Range | Excel.RangeReference | Excel.FunctionRes...",
        examples: [],
      },
      {
        name: "Excel.Functions.vlookup",
        description:
          "Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify. By default, the table must be sorted in an ascending order.",
        kind: "Method",
        signature:
          "Excel.Functions.vlookup(lookupValue: string | number | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, tableArray: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<...>, colIndexNum: number | ... 2 more ... | Excel.FunctionResult<...>, rangeLookup?: boolean | ... 2 more ... | Excel.FunctionResul...",
        examples: ['let unitSoldInNov = workbook.functions.vlookup("Wrench", range, 2, false);'],
      },
      {
        name: "Excel.Functions.weekday",
        description: "Returns a number from 1 to 7 identifying the day of the week of a date.",
        kind: "Method",
        signature:
          "Excel.Functions.weekday => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, returnType?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.weekNum",
        description: "Returns the week number in the year.",
        kind: "Method",
        signature:
          "Excel.Functions.weekNum => (serialNumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, returnType?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.weibull_Dist",
        description: "Returns the Weibull distribution.",
        kind: "Method",
        signature:
          "Excel.Functions.weibull_Dist => (x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<a...",
        examples: [],
      },
      {
        name: "Excel.Functions.workDay",
        description:
          "Returns the serial number of the date before or after a specified number of workdays.",
        kind: "Method",
        signature:
          "Excel.Functions.workDay => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, days: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionR...",
        examples: [],
      },
      {
        name: "Excel.Functions.workDay_Intl",
        description:
          "Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.",
        kind: "Method",
        signature:
          "Excel.Functions.workDay_Intl => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, days: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, weekend?: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | str...",
        examples: [],
      },
      {
        name: "Excel.Functions.xirr",
        description: "Returns the internal rate of return for a schedule of cash flows.",
        kind: "Method",
        signature:
          "Excel.Functions.xirr => (values: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>, dates: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>, guess?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult...",
        examples: [],
      },
      {
        name: "Excel.Functions.xnpv",
        description: "Returns the net present value for a schedule of cash flows.",
        kind: "Method",
        signature:
          "Excel.Functions.xnpv => (rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, values: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>, dates: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<n...",
        examples: [],
      },
      {
        name: "Excel.Functions.xor",
        description: "Returns a logical 'Exclusive Or' of all arguments.",
        kind: "Method",
        signature:
          "Excel.Functions.xor => (...values: Array<boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>) => FunctionResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Functions.year",
        description: "Returns the year of a date, an integer in the range 1900 - 9999.",
        kind: "Method",
        signature:
          "Excel.Functions.year => (serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
      {
        name: "Excel.Functions.yearFrac",
        description:
          "Returns the year fraction representing the number of whole days between start_date and end_date.",
        kind: "Method",
        signature:
          "Excel.Functions.yearFrac => (startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionR...",
        examples: [],
      },
      {
        name: "Excel.Functions.yield",
        description: "Returns the yield on a security that pays periodic interest.",
        kind: "Method",
        signature:
          "Excel.Functions.yield => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number |...",
        examples: [],
      },
      {
        name: "Excel.Functions.yieldDisc",
        description:
          "Returns the annual yield for a discounted security. For example, a treasury bill.",
        kind: "Method",
        signature:
          "Excel.Functions.yieldDisc => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: nu...",
        examples: [],
      },
      {
        name: "Excel.Functions.yieldMat",
        description: "Returns the annual yield of a security that pays interest at maturity.",
        kind: "Method",
        signature:
          "Excel.Functions.yieldMat => (settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: numbe...",
        examples: [],
      },
      {
        name: "Excel.Functions.z_Test",
        description: "Returns the one-tailed P-value of a z-test.",
        kind: "Method",
        signature:
          "Excel.Functions.z_Test => (array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, sigma?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>) => FunctionResult<number>",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.GeometricShape",
    apiList: [
      {
        name: "Excel.GeometricShape.id",
        description: "Returns the shape identifier.",
        kind: "Property",
        signature: "Excel.GeometricShape.id: string",
        examples: [],
      },
      {
        name: "Excel.GeometricShape.shape",
        description: "Returns the `Shape` object for the geometric shape.",
        kind: "Property",
        signature: "Excel.GeometricShape.shape: Shape",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.GettingDataErrorCellValue",
    apiList: [
      {
        name: "Excel.GettingDataErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.GettingDataErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.GettingDataErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.GettingDataErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.GettingDataErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.GettingDataErrorCellValue.errorType: ErrorCellValueType.gettingData | "GettingData"',
        examples: [],
      },
      {
        name: "Excel.GettingDataErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.GettingDataErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.GetUsedRangeAreasOptions",
    apiList: [
      {
        name: "Excel.GetUsedRangeAreasOptions.excludeNamedRanges",
        description:
          "If true, then range areas that are entirely a single named range are excluded. Range areas that include a names range and other contiguous data are always returned. By default, named ranges are not excluded.",
        kind: "Property",
        signature: "Excel.GetUsedRangeAreasOptions.excludeNamedRanges: boolean",
        examples: [],
      },
      {
        name: "Excel.GetUsedRangeAreasOptions.excludePivotTables",
        description:
          "If true, then range areas that are entirely a single PivotTable are excluded. Range areas that include a PivotTable and other contiguous data are always returned. By default, PivotTables are not excluded.",
        kind: "Property",
        signature: "Excel.GetUsedRangeAreasOptions.excludePivotTables: boolean",
        examples: [],
      },
      {
        name: "Excel.GetUsedRangeAreasOptions.excludeTables",
        description:
          "If true, then range areas that are entirely a single table are excluded. Range areas that include a table and other contiguous data are always returned. By default, tables are not excluded.",
        kind: "Property",
        signature: "Excel.GetUsedRangeAreasOptions.excludeTables: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.GroupShapeCollection",
    apiList: [
      {
        name: "Excel.GroupShapeCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.GroupShapeCollection.items: Shape[]",
        examples: [],
      },
      {
        name: "Excel.GroupShapeCollection.getCount",
        description: "Returns the number of shapes in the shape group.",
        kind: "Method",
        signature:
          "Excel.GroupShapeCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.GroupShapeCollection.getItem",
        description: "Gets a shape using its name or ID.",
        kind: "Method",
        signature: "Excel.GroupShapeCollection.getItem => (key: string) => Excel.Shape",
        examples: [],
      },
      {
        name: "Excel.GroupShapeCollection.getItemAt",
        description: "Gets a shape based on its position in the collection.",
        kind: "Method",
        signature: "Excel.GroupShapeCollection.getItemAt => (index: number) => Excel.Shape",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.GuidedReapplyManager",
    apiList: [
      {
        name: "Excel.GuidedReapplyManager.activities",
        description:
          "The `UserActivity` list that represents user changes which did not upload successfully into the document. Data is only valid after a call to `updateActivities`.",
        kind: "Property",
        signature: "Excel.GuidedReapplyManager.activities: UserActivityCollection",
        examples: [],
      },
      {
        name: "Excel.GuidedReapplyManager.discardActivites",
        description: "Discards all guided reapply content.",
        kind: "Method",
        signature: "Excel.GuidedReapplyManager.discardActivites => () => void",
        examples: [],
      },
      {
        name: "Excel.GuidedReapplyManager.openSavedFile",
        description:
          "Opens the saved workbook file in read-only mode. This file is created after a user encounters a coauthoring error and reloads the document.",
        kind: "Method",
        signature: "Excel.GuidedReapplyManager.openSavedFile => () => void",
        examples: [],
      },
      {
        name: "Excel.GuidedReapplyManager.reapplyActivity",
        description: "Adds the activity back into the workbook after a coauthoring error.",
        kind: "Method",
        signature:
          "Excel.GuidedReapplyManager.reapplyActivity => (activity: Excel.UserActivity) => void",
        examples: [],
      },
      {
        name: "Excel.GuidedReapplyManager.saveActivities",
        description:
          "Saves a copy of guided reapply content for comparing against the server version of the workbook.",
        kind: "Method",
        signature: "Excel.GuidedReapplyManager.saveActivities => () => void",
        examples: [],
      },
      {
        name: "Excel.GuidedReapplyManager.updateActivities",
        description:
          "A call to update the `UserActivity` list from the guided reapply data. Called when new content is available for the activities collection.",
        kind: "Method",
        signature: "Excel.GuidedReapplyManager.updateActivities => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.HeaderFooter",
    apiList: [
      {
        name: "Excel.HeaderFooter.centerFooter",
        description:
          "The center footer of the worksheet. To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.",
        kind: "Property",
        signature: "Excel.HeaderFooter.centerFooter: string",
        examples: [],
      },
      {
        name: "Excel.HeaderFooter.centerHeader",
        description:
          "The center header of the worksheet. To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.",
        kind: "Property",
        signature: "Excel.HeaderFooter.centerHeader: string",
        examples: [],
      },
      {
        name: "Excel.HeaderFooter.leftFooter",
        description:
          "The left footer of the worksheet. To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.",
        kind: "Property",
        signature: "Excel.HeaderFooter.leftFooter: string",
        examples: [],
      },
      {
        name: "Excel.HeaderFooter.leftHeader",
        description:
          "The left header of the worksheet. To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.",
        kind: "Property",
        signature: "Excel.HeaderFooter.leftHeader: string",
        examples: [],
      },
      {
        name: "Excel.HeaderFooter.rightFooter",
        description:
          "The right footer of the worksheet. To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.",
        kind: "Property",
        signature: "Excel.HeaderFooter.rightFooter: string",
        examples: [],
      },
      {
        name: "Excel.HeaderFooter.rightHeader",
        description:
          "The right header of the worksheet. To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.",
        kind: "Property",
        signature: "Excel.HeaderFooter.rightHeader: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.HeaderFooterGroup",
    apiList: [
      {
        name: "Excel.HeaderFooterGroup.defaultForAllPages",
        description:
          "The general header/footer, used for all pages unless even/odd or first page is specified.",
        kind: "Property",
        signature: "Excel.HeaderFooterGroup.defaultForAllPages: HeaderFooter",
        examples: [],
      },
      {
        name: "Excel.HeaderFooterGroup.evenPages",
        description:
          "The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.",
        kind: "Property",
        signature: "Excel.HeaderFooterGroup.evenPages: HeaderFooter",
        examples: [],
      },
      {
        name: "Excel.HeaderFooterGroup.firstPage",
        description:
          "The first page header/footer, for all other pages general or even/odd is used.",
        kind: "Property",
        signature: "Excel.HeaderFooterGroup.firstPage: HeaderFooter",
        examples: [],
      },
      {
        name: "Excel.HeaderFooterGroup.oddPages",
        description:
          "The header/footer to use for odd pages, even header/footer needs to be specified for even pages.",
        kind: "Property",
        signature: "Excel.HeaderFooterGroup.oddPages: HeaderFooter",
        examples: [],
      },
      {
        name: "Excel.HeaderFooterGroup.state",
        description:
          "The state by which headers/footers are set. See `Excel.HeaderFooterState` for details.",
        kind: "Property",
        signature:
          'Excel.HeaderFooterGroup.state: HeaderFooterState | "Default" | "FirstAndDefault" | "OddAndEven" | "FirstOddAndEven"',
        examples: [],
      },
      {
        name: "Excel.HeaderFooterGroup.useSheetMargins",
        description:
          "Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.",
        kind: "Property",
        signature: "Excel.HeaderFooterGroup.useSheetMargins: boolean",
        examples: [],
      },
      {
        name: "Excel.HeaderFooterGroup.useSheetScale",
        description:
          "Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.",
        kind: "Property",
        signature: "Excel.HeaderFooterGroup.useSheetScale: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Icon",
    apiList: [
      {
        name: "Excel.Icon.index",
        description: "Specifies the index of the icon in the given set.",
        kind: "Property",
        signature: "Excel.Icon.index: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.IconSetConditionalFormat",
    apiList: [
      {
        name: "Excel.IconSetConditionalFormat.criteria",
        description:
          "An array of criteria and icon sets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.",
        kind: "Property",
        signature: "Excel.IconSetConditionalFormat.criteria: Excel.ConditionalIconCriterion[]",
        examples: [],
      },
      {
        name: "Excel.IconSetConditionalFormat.reverseIconOrder",
        description:
          "If `true`, reverses the icon orders for the icon set. Note that this cannot be set if custom icons are used.",
        kind: "Property",
        signature: "Excel.IconSetConditionalFormat.reverseIconOrder: boolean",
        examples: [],
      },
      {
        name: "Excel.IconSetConditionalFormat.showIconOnly",
        description: "If `true`, hides the values and only shows icons.",
        kind: "Property",
        signature: "Excel.IconSetConditionalFormat.showIconOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.IconSetConditionalFormat.style",
        description: "If set, displays the icon set option for the conditional format.",
        kind: "Property",
        signature:
          'Excel.IconSetConditionalFormat.style: Excel.IconSet | "Invalid" | "ThreeArrows" | "ThreeArrowsGray" | "ThreeFlags" | "ThreeTrafficLights1" | "ThreeTrafficLights2" | "ThreeSigns" | "ThreeSymbols" | "ThreeSymbols2" | ... 11 more ... | "FiveBoxes"',
        examples: [
          "conditionalFormat.iconSetOrNullObject.style = Excel.IconSet.fourTrafficLights;",
        ],
      },
    ],
  },
  {
    objName: "Excel.Identity",
    apiList: [
      {
        name: "Excel.Identity.displayName",
        description: "Represents the user's display name.",
        kind: "Property",
        signature: "Excel.Identity.displayName: string",
        examples: [],
      },
      {
        name: "Excel.Identity.id",
        description: "Represents the user's unique ID.",
        kind: "Property",
        signature: "Excel.Identity.id: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Image",
    apiList: [
      {
        name: "Excel.Image.format",
        description: "Returns the format of the image.",
        kind: "Property",
        signature:
          'Excel.Image.format: PictureFormat | "UNKNOWN" | "BMP" | "JPEG" | "GIF" | "PNG" | "SVG"',
        examples: ['"The image\'s format is: " + image.format;'],
      },
      {
        name: "Excel.Image.id",
        description: "Specifies the shape identifier for the image object.",
        kind: "Property",
        signature: "Excel.Image.id: string",
        examples: [],
      },
      {
        name: "Excel.Image.shape",
        description: "Returns the `Shape` object associated with the image.",
        kind: "Property",
        signature: "Excel.Image.shape: Shape",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.InsertWorksheetOptions",
    apiList: [
      {
        name: "Excel.InsertWorksheetOptions.positionType",
        description:
          'The insert position, in the current workbook, of the new worksheets. See `Excel.WorksheetPositionType` for details. The default position is "End".',
        kind: "Property",
        signature:
          'Excel.InsertWorksheetOptions.positionType: "None" | "Before" | "After" | WorksheetPositionType | "Beginning" | "End"',
        examples: [],
      },
      {
        name: "Excel.InsertWorksheetOptions.relativeTo",
        description:
          "The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter. The default is `null`. If the `relativeTo` parameter is not set, worksheets will be inserted based on `positionType`, at the start or end of the current workbook.",
        kind: "Property",
        signature: "Excel.InsertWorksheetOptions.relativeTo: string | Worksheet",
        examples: [],
      },
      {
        name: "Excel.InsertWorksheetOptions.sheetNamesToInsert",
        description:
          "The names of individual worksheets to insert. By default, all the worksheets from the source workbook are inserted.",
        kind: "Property",
        signature: "Excel.InsertWorksheetOptions.sheetNamesToInsert: string[]",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.IterativeCalculation",
    apiList: [
      {
        name: "Excel.IterativeCalculation.enabled",
        description: "True if Excel will use iteration to resolve circular references.",
        kind: "Property",
        signature: "Excel.IterativeCalculation.enabled: boolean",
        examples: [],
      },
      {
        name: "Excel.IterativeCalculation.maxChange",
        description:
          "Specifies the maximum amount of change between each iteration as Excel resolves circular references.",
        kind: "Property",
        signature: "Excel.IterativeCalculation.maxChange: number",
        examples: [],
      },
      {
        name: "Excel.IterativeCalculation.maxIteration",
        description:
          "Specifies the maximum number of iterations that Excel can use to resolve a circular reference.",
        kind: "Property",
        signature: "Excel.IterativeCalculation.maxIteration: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Line",
    apiList: [
      {
        name: "Excel.Line.beginArrowheadLength",
        description:
          "Represents the length of the arrowhead at the beginning of the specified line.",
        kind: "Property",
        signature:
          'Excel.Line.beginArrowheadLength: Excel.ArrowheadLength | "Short" | "Medium" | "Long"',
        examples: ["line.beginArrowheadLength = Excel.ArrowheadLength.long;"],
      },
      {
        name: "Excel.Line.beginArrowheadStyle",
        description:
          "Represents the style of the arrowhead at the beginning of the specified line.",
        kind: "Property",
        signature:
          'Excel.Line.beginArrowheadStyle: Excel.ArrowheadStyle | "None" | "Triangle" | "Stealth" | "Diamond" | "Oval" | "Open"',
        examples: ["line.beginArrowheadStyle = Excel.ArrowheadStyle.oval;"],
      },
      {
        name: "Excel.Line.beginArrowheadWidth",
        description:
          "Represents the width of the arrowhead at the beginning of the specified line.",
        kind: "Property",
        signature:
          'Excel.Line.beginArrowheadWidth: "Medium" | Excel.ArrowheadWidth | "Narrow" | "Wide"',
        examples: ["line.beginArrowheadWidth = Excel.ArrowheadWidth.wide;"],
      },
      {
        name: "Excel.Line.beginConnectedShape",
        description:
          "Represents the shape to which the beginning of the specified line is attached.",
        kind: "Property",
        signature: "Excel.Line.beginConnectedShape: Shape",
        examples: [],
      },
      {
        name: "Excel.Line.beginConnectedSite",
        description:
          "Represents the connection site to which the beginning of a connector is connected. Returns `null` when the beginning of the line is not attached to any shape.",
        kind: "Property",
        signature: "Excel.Line.beginConnectedSite: number",
        examples: [],
      },
      {
        name: "Excel.Line.connectorType",
        description: "Represents the connector type for the line.",
        kind: "Property",
        signature: 'Excel.Line.connectorType: ConnectorType | "Straight" | "Elbow" | "Curve"',
        examples: [],
      },
      {
        name: "Excel.Line.endArrowheadLength",
        description: "Represents the length of the arrowhead at the end of the specified line.",
        kind: "Property",
        signature:
          'Excel.Line.endArrowheadLength: Excel.ArrowheadLength | "Short" | "Medium" | "Long"',
        examples: ["line.endArrowheadLength = Excel.ArrowheadLength.long;"],
      },
      {
        name: "Excel.Line.endArrowheadStyle",
        description: "Represents the style of the arrowhead at the end of the specified line.",
        kind: "Property",
        signature:
          'Excel.Line.endArrowheadStyle: Excel.ArrowheadStyle | "None" | "Triangle" | "Stealth" | "Diamond" | "Oval" | "Open"',
        examples: ["line.endArrowheadStyle = Excel.ArrowheadStyle.triangle;"],
      },
      {
        name: "Excel.Line.endArrowheadWidth",
        description: "Represents the width of the arrowhead at the end of the specified line.",
        kind: "Property",
        signature:
          'Excel.Line.endArrowheadWidth: "Medium" | Excel.ArrowheadWidth | "Narrow" | "Wide"',
        examples: ["line.endArrowheadWidth = Excel.ArrowheadWidth.wide;"],
      },
      {
        name: "Excel.Line.endConnectedShape",
        description: "Represents the shape to which the end of the specified line is attached.",
        kind: "Property",
        signature: "Excel.Line.endConnectedShape: Shape",
        examples: [],
      },
      {
        name: "Excel.Line.endConnectedSite",
        description:
          "Represents the connection site to which the end of a connector is connected. Returns `null` when the end of the line is not attached to any shape.",
        kind: "Property",
        signature: "Excel.Line.endConnectedSite: number",
        examples: [],
      },
      {
        name: "Excel.Line.id",
        description: "Specifies the shape identifier.",
        kind: "Property",
        signature: "Excel.Line.id: string",
        examples: [],
      },
      {
        name: "Excel.Line.isBeginConnected",
        description: "Specifies if the beginning of the specified line is connected to a shape.",
        kind: "Property",
        signature: "Excel.Line.isBeginConnected: boolean",
        examples: [],
      },
      {
        name: "Excel.Line.isEndConnected",
        description: "Specifies if the end of the specified line is connected to a shape.",
        kind: "Property",
        signature: "Excel.Line.isEndConnected: boolean",
        examples: [],
      },
      {
        name: "Excel.Line.shape",
        description: "Returns the `Shape` object associated with the line.",
        kind: "Property",
        signature: "Excel.Line.shape: Shape",
        examples: [],
      },
      {
        name: "Excel.Line.connectBeginShape",
        description: "Attaches the beginning of the specified connector to a specified shape.",
        kind: "Method",
        signature:
          "Excel.Line.connectBeginShape(shape: Excel.Shape, connectionSite: number) => void",
        examples: ['line.connectBeginShape(shapes.getItem("Left"), 2);'],
      },
      {
        name: "Excel.Line.connectEndShape",
        description: "Attaches the end of the specified connector to a specified shape.",
        kind: "Method",
        signature: "Excel.Line.connectEndShape(shape: Excel.Shape, connectionSite: number) => void",
        examples: ['line.connectEndShape(shapes.getItem("Right"), 0);'],
      },
      {
        name: "Excel.Line.disconnectBeginShape",
        description: "Detaches the beginning of the specified connector from a shape.",
        kind: "Method",
        signature: "Excel.Line.disconnectBeginShape() => void",
        examples: ["line.disconnectBeginShape();"],
      },
      {
        name: "Excel.Line.disconnectEndShape",
        description: "Detaches the end of the specified connector from a shape.",
        kind: "Method",
        signature: "Excel.Line.disconnectEndShape() => void",
        examples: ["line.disconnectEndShape();"],
      },
    ],
  },
  {
    objName: "Excel.LineageActivityCollection",
    apiList: [
      {
        name: "Excel.LineageActivityCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.LineageActivityCollection.items: UserActivity[]",
        examples: [],
      },
      {
        name: "Excel.LineageActivityCollection.clear",
        description: "Clears all loaded activities and resets filter data.",
        kind: "Method",
        signature: "Excel.LineageActivityCollection.clear => () => void",
        examples: [],
      },
      {
        name: "Excel.LineageActivityCollection.getCount",
        description: "Gets the number of activities in the collection.",
        kind: "Method",
        signature:
          "Excel.LineageActivityCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.LineageActivityCollection.getItemAt",
        description: "Gets the UserActivity object by its index in the collection.",
        kind: "Method",
        signature:
          "Excel.LineageActivityCollection.getItemAt => (index: number) => Excel.UserActivity",
        examples: [],
      },
      {
        name: "Excel.LineageActivityCollection.getState",
        description: "Gets the current lineage state after loading activities.",
        kind: "Method",
        signature:
          "Excel.LineageActivityCollection.getState => () => OfficeExtension.ClientResult<Excel.LineageState>",
        examples: [],
      },
      {
        name: "Excel.LineageActivityCollection.updateActivities",
        description:
          "Updates stale activities. This applies the current filter and indicates if there are new activities. Should be called after the activityUpdate event is raised.",
        kind: "Method",
        signature: "Excel.LineageActivityCollection.updateActivities => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.LineageOptions",
    apiList: [
      {
        name: "Excel.LineageOptions.capacity",
        description: "Represents the requested capacity from client.",
        kind: "Property",
        signature: "Excel.LineageOptions.capacity: number",
        examples: [],
      },
      {
        name: "Excel.LineageOptions.filter",
        description:
          "Represents the filter information that will be applied when loading activities.",
        kind: "Property",
        signature: "Excel.LineageOptions.filter: UserActivityFilter",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.LineageState",
    apiList: [
      {
        name: "Excel.LineageState.correlationId",
        description:
          "Unique correlation ID representing the Excel client's end of log state after each load operation.",
        kind: "Property",
        signature: "Excel.LineageState.correlationId: string",
        examples: [],
      },
      {
        name: "Excel.LineageState.endOfLogStatus",
        description: "Represents the state of the revision log after loading activities.",
        kind: "Property",
        signature:
          'Excel.LineageState.endOfLogStatus: "Error" | LineageEndOfLogStatus | "LoadInProgress" | "Success" | "EndOfLog" | "Purged" | "Trimmed" | "Unsupported" | "Cleared"',
        examples: [],
      },
      {
        name: "Excel.LineageState.filter",
        description:
          "Represents the filter information that will be applied when loading Lineage activities.",
        kind: "Property",
        signature: "Excel.LineageState.filter: UserActivityFilter",
        examples: [],
      },
      {
        name: "Excel.LineageState.firstViewActivityId",
        description:
          "First activity's ID stored in the Excel client. Activities with activityId < firstViewActivityId should be removed to keep them in sync with the Excel client.",
        kind: "Property",
        signature: "Excel.LineageState.firstViewActivityId: number",
        examples: [],
      },
      {
        name: "Excel.LineageState.historyClearedBy",
        description:
          "The author who cleared previous activities. This is set when endOfLogStatus is Cleared.",
        kind: "Property",
        signature: "Excel.LineageState.historyClearedBy: Identity",
        examples: [],
      },
      {
        name: "Excel.LineageState.historyClearedByAuthorEmail",
        description:
          "Email of the author who cleared the previous activities. This is set when endOfLogStatus is Cleared.",
        kind: "Property",
        signature: "Excel.LineageState.historyClearedByAuthorEmail: string",
        examples: [],
      },
      {
        name: "Excel.LineageState.historyClearedDateTime",
        description:
          "The date at which previous activities were cleared. This is set when endOfLogStatus is Cleared.",
        kind: "Property",
        signature: "Excel.LineageState.historyClearedDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.LineageState.lastSearchedDateTime",
        description:
          "The date of the last searched activity that the LineageActivityCollection has completed processing. This can be different from the date of the activity collection's last item.",
        kind: "Property",
        signature: "Excel.LineageState.lastSearchedDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.LineageState.lastViewActivityId",
        description:
          "Last activity's ID stored in the Excel client. Activities with activityId > lastViewActivityId should be removed to keep them in sync with the Excel client.",
        kind: "Property",
        signature: "Excel.LineageState.lastViewActivityId: number",
        examples: [],
      },
      {
        name: "Excel.LineageState.newActivitiesAvailable",
        description: "Indicates if there are newer activities to be processed.",
        kind: "Property",
        signature: "Excel.LineageState.newActivitiesAvailable: boolean",
        examples: [],
      },
      {
        name: "Excel.LineageState.previousActivitiesAvailable",
        description:
          "Flag indicating if additional activities with activityId > lastViewActivityId are available to load.",
        kind: "Property",
        signature: "Excel.LineageState.previousActivitiesAvailable: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.LinkedEntityCellValue",
    apiList: [
      {
        name: "Excel.LinkedEntityCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.LinkedEntityCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.LinkedEntityCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.cardLayout",
        description:
          'Represents the layout of this linked entity in card view. If the `CardLayout` object doesn\'t have a layout property, it default value is "Entity".',
        kind: "Property",
        signature: "Excel.LinkedEntityCellValue.cardLayout: CardLayout",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.id",
        description: "Represents the service source that provided the information in this value.",
        kind: "Property",
        signature: "Excel.LinkedEntityCellValue.id: LinkedEntityId",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.properties",
        description: "Represents the properties of this linked entity and their metadata.",
        kind: "Property",
        signature:
          "Excel.LinkedEntityCellValue.properties: { [key: string]: CellValue & { propertyMetadata?: CellValuePropertyMetadata; }; }",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.provider",
        description:
          "Represents information that describes the service that provided data in this `LinkedEntityCellValue`. This information can be used for branding in entity cards.",
        kind: "Property",
        signature: "Excel.LinkedEntityCellValue.provider: CellValueProviderAttributes",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.text",
        description: "Represents the text shown when a cell with this value is rendered.",
        kind: "Property",
        signature: "Excel.LinkedEntityCellValue.text: string",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.LinkedEntityCellValue.type: CellValueType.linkedEntity | "LinkedEntity"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.LinkedEntityId",
    apiList: [
      {
        name: "Excel.LinkedEntityId.culture",
        description: "Represents which language culture was used to create this `CellValue`.",
        kind: "Property",
        signature: "Excel.LinkedEntityId.culture: string",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityId.domainId",
        description: "Represents a domain specific to a service used to create the `CellValue`.",
        kind: "Property",
        signature: "Excel.LinkedEntityId.domainId: string",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityId.entityId",
        description:
          "Represents an identifier specific to a service used to create the `CellValue`.",
        kind: "Property",
        signature: "Excel.LinkedEntityId.entityId: string",
        examples: [],
      },
      {
        name: "Excel.LinkedEntityId.serviceId",
        description: "Represents which service was used to create the `CellValue`.",
        kind: "Property",
        signature: "Excel.LinkedEntityId.serviceId: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ListDataValidation",
    apiList: [
      {
        name: "Excel.ListDataValidation.inCellDropDown",
        description:
          "Specifies whether to display the list in a cell drop-down. The default is `true`.",
        kind: "Property",
        signature: "Excel.ListDataValidation.inCellDropDown: boolean",
        examples: [],
      },
      {
        name: "Excel.ListDataValidation.source",
        description:
          "Source of the list for data validation When setting the value, it can be passed in as a `Range` object, or a string that contains a comma-separated number, boolean, or date.",
        kind: "Property",
        signature: "Excel.ListDataValidation.source: string | Range",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NamedItem",
    apiList: [
      {
        name: "Excel.NamedItem.arrayValues",
        description: "Returns an object containing values and types of the named item.",
        kind: "Property",
        signature: "Excel.NamedItem.arrayValues: NamedItemArrayValues",
        examples: [],
      },
      {
        name: "Excel.NamedItem.comment",
        description: "Specifies the comment associated with this name.",
        kind: "Property",
        signature: "Excel.NamedItem.comment: string",
        examples: [],
      },
      {
        name: "Excel.NamedItem.formula",
        description:
          'The formula of the named item. Formulas always start with an equal sign ("=").',
        kind: "Property",
        signature: "Excel.NamedItem.formula: any",
        examples: [
          'myNamedItem.formula = "=Sample!$B$10:$D$14";',
          '`Just updated the named item "${myNamedItem.name}" -- it\'s now located here: ${myNamedItem.formula}`;',
        ],
      },
      {
        name: "Excel.NamedItem.name",
        description: "The name of the object.",
        kind: "Property",
        signature: "Excel.NamedItem.name: string",
        examples: [
          '`Just updated the named item "${myNamedItem.name}" -- it\'s now located here: ${myNamedItem.formula}`;',
        ],
      },
      {
        name: "Excel.NamedItem.scope",
        description:
          "Specifies if the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook.",
        kind: "Property",
        signature: 'Excel.NamedItem.scope: NamedItemScope | "Worksheet" | "Workbook"',
        examples: [],
      },
      {
        name: "Excel.NamedItem.type",
        description:
          "Specifies the type of the value returned by the name's formula. See `Excel.NamedItemType` for details.",
        kind: "Property",
        signature:
          'Excel.NamedItem.type: Excel.NamedItemType | "String" | "Integer" | "Double" | "Boolean" | "Range" | "Error" | "Array"',
        examples: ["namedItem.type;", "nameditem.type;"],
      },
      {
        name: "Excel.NamedItem.value",
        description:
          "Represents the value computed by the name's formula. For a named range, will return the range address.",
        kind: "Property",
        signature: "Excel.NamedItem.value: any",
        examples: [],
      },
      {
        name: "Excel.NamedItem.valueAsJson",
        description:
          "A JSON representation of the values in this named item. Unlike `NamedItem.value`, `NamedItem.valueAsJson` supports all data types which can be in a cell. Examples include formatted number values and web images, in addition to the standard boolean, number, and string values. Data returned from this API always aligns with the en-US locale. To retrieve data in the user's display locale, use `NamedItem.valueAsJsonLocal`.",
        kind: "Property",
        signature: "Excel.NamedItem.valueAsJson: string | CellValue",
        examples: [],
      },
      {
        name: "Excel.NamedItem.valueAsJsonLocal",
        description:
          "A JSON representation of the values in this named item. Unlike `NamedItem.value`, `NamedItem.valueAsJsonLocal` supports all data types which can be in a cell. Examples include formatted number values and web images, in addition to the standard boolean, number, and string values. Data returned from this API always aligns with the user's display locale. To retrieve data independent of locale, use `NamedItem.valueAsJson`.",
        kind: "Property",
        signature: "Excel.NamedItem.valueAsJsonLocal: string | CellValue",
        examples: [],
      },
      {
        name: "Excel.NamedItem.visible",
        description: "Specifies if the object is visible.",
        kind: "Property",
        signature: "Excel.NamedItem.visible: boolean",
        examples: [],
      },
      {
        name: "Excel.NamedItem.worksheet",
        description:
          "Returns the worksheet on which the named item is scoped to. Throws an error if the item is scoped to the workbook instead.",
        kind: "Property",
        signature: "Excel.NamedItem.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.NamedItem.worksheetOrNullObject",
        description:
          "Returns the worksheet to which the named item is scoped. If the item is scoped to the workbook instead, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Property",
        signature: "Excel.NamedItem.worksheetOrNullObject: Worksheet",
        examples: [],
      },
      {
        name: "Excel.NamedItem.delete",
        description: "Deletes the given name.",
        kind: "Method",
        signature: "Excel.NamedItem.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.NamedItem.getRange",
        description:
          "Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.",
        kind: "Method",
        signature: "Excel.NamedItem.getRange() => Excel.Range",
        examples: ['const range = names.getItem("MyRange").getRange();'],
      },
      {
        name: "Excel.NamedItem.getRangeOrNullObject",
        description:
          "Returns the range object that is associated with the name. If the named item's type is not a range, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.NamedItem.getRangeOrNullObject => () => Excel.Range",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NamedItemArrayValues",
    apiList: [
      {
        name: "Excel.NamedItemArrayValues.types",
        description: "Represents the types for each item in the named item array",
        kind: "Property",
        signature: "Excel.NamedItemArrayValues.types: RangeValueType[][]",
        examples: [],
      },
      {
        name: "Excel.NamedItemArrayValues.values",
        description: "Represents the values of each item in the named item array.",
        kind: "Property",
        signature: "Excel.NamedItemArrayValues.values: any[][]",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NamedItemCollection",
    apiList: [
      {
        name: "Excel.NamedItemCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.NamedItemCollection.items: NamedItem[]",
        examples: [],
      },
      {
        name: "Excel.NamedItemCollection.add",
        description: "Adds a new name to the collection of the given scope.",
        kind: "Method",
        signature:
          "Excel.NamedItemCollection.add(name: string, reference: string | Excel.Range, comment?: string) => Excel.NamedItem",
        examples: ['activeWorksheet.names.add("ExpensesHeader", headerRange);'],
      },
      {
        name: "Excel.NamedItemCollection.addFormulaLocal",
        description:
          "Adds a new name to the collection of the given scope using the user's locale for the formula.",
        kind: "Method",
        signature:
          "Excel.NamedItemCollection.addFormulaLocal => (name: string, formula: string, comment?: string) => Excel.NamedItem",
        examples: [],
      },
      {
        name: "Excel.NamedItemCollection.getCount",
        description: "Gets the number of named items in the collection.",
        kind: "Method",
        signature:
          "Excel.NamedItemCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.NamedItemCollection.getItem",
        description: "Gets a `NamedItem` object using its name.",
        kind: "Method",
        signature: "Excel.NamedItemCollection.getItem(name: string) => Excel.NamedItem",
        examples: [
          'const range = names.getItem("MyRange").getRange();',
          'const namedItem = names.getItem("MyRange");',
          "const nameditem = workbook.names.getItem(sheetName);",
        ],
      },
    ],
  },
  {
    objName: "Excel.NamedSheetView",
    apiList: [
      {
        name: "Excel.NamedSheetView.name",
        description:
          'Gets or sets the name of the sheet view. The temporary sheet view name is the empty string (""). Naming the view by using the name property causes the sheet view to be saved.',
        kind: "Property",
        signature: "Excel.NamedSheetView.name: string",
        examples: [],
      },
      {
        name: "Excel.NamedSheetView.activate",
        description:
          'Activates this sheet view. This is equivalent to using "Switch To" in the Excel UI.',
        kind: "Method",
        signature: "Excel.NamedSheetView.activate => () => void",
        examples: [],
      },
      {
        name: "Excel.NamedSheetView.delete",
        description: "Removes the sheet view from the worksheet.",
        kind: "Method",
        signature: "Excel.NamedSheetView.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.NamedSheetView.duplicate",
        description: "Creates a copy of this sheet view.",
        kind: "Method",
        signature: "Excel.NamedSheetView.duplicate => (name?: string) => Excel.NamedSheetView",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NamedSheetViewCollection",
    apiList: [
      {
        name: "Excel.NamedSheetViewCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.NamedSheetViewCollection.items: NamedSheetView[]",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.add",
        description: "Creates a new sheet view with the given name.",
        kind: "Method",
        signature: "Excel.NamedSheetViewCollection.add => (name: string) => Excel.NamedSheetView",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.enterTemporary",
        description:
          'Creates and activates a new temporary sheet view. Temporary views are removed when closing the application, exiting the temporary view with the exit method, or switching to another sheet view. The temporary sheet view can also be acccessed with the empty string (""), if the temporary view exists.',
        kind: "Method",
        signature: "Excel.NamedSheetViewCollection.enterTemporary => () => Excel.NamedSheetView",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.exit",
        description: "Exits the currently active sheet view.",
        kind: "Method",
        signature: "Excel.NamedSheetViewCollection.exit => () => void",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.getActive",
        description: "Gets the worksheet's currently active sheet view.",
        kind: "Method",
        signature: "Excel.NamedSheetViewCollection.getActive => () => Excel.NamedSheetView",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.getCount",
        description:
          "Gets the number of sheet views in this worksheet. Includes the temporary sheet view if it exists.",
        kind: "Method",
        signature:
          "Excel.NamedSheetViewCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.getItem",
        description: "Gets a sheet view using its name.",
        kind: "Method",
        signature:
          "Excel.NamedSheetViewCollection.getItem => (key: string) => Excel.NamedSheetView",
        examples: [],
      },
      {
        name: "Excel.NamedSheetViewCollection.getItemAt",
        description: "Gets a sheet view by its index in the collection.",
        kind: "Method",
        signature:
          "Excel.NamedSheetViewCollection.getItemAt => (index: number) => Excel.NamedSheetView",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NameErrorCellValue",
    apiList: [
      {
        name: "Excel.NameErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.NameErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.NameErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.NameErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.NameErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.NameErrorCellValue.errorType: ErrorCellValueType.name | "Name"',
        examples: [],
      },
      {
        name: "Excel.NameErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.NameErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NotAvailableErrorCellValue",
    apiList: [
      {
        name: "Excel.NotAvailableErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.NotAvailableErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.NotAvailableErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.NotAvailableErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.NotAvailableErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.NotAvailableErrorCellValue.errorType: ErrorCellValueType.notAvailable | "NotAvailable"',
        examples: [],
      },
      {
        name: "Excel.NotAvailableErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.NotAvailableErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NullErrorCellValue",
    apiList: [
      {
        name: "Excel.NullErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.NullErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.NullErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.NullErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.NullErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.NullErrorCellValue.errorType: ErrorCellValueType.null | "Null"',
        examples: [],
      },
      {
        name: "Excel.NullErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.NullErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NumberFormatInfo",
    apiList: [
      {
        name: "Excel.NumberFormatInfo.currencySymbol",
        description:
          "Gets the currency symbol for currency values. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.NumberFormatInfo.currencySymbol: string",
        examples: [],
      },
      {
        name: "Excel.NumberFormatInfo.numberDecimalSeparator",
        description:
          "Gets the string used as the decimal separator for numeric values. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.NumberFormatInfo.numberDecimalSeparator: string",
        examples: [
          "const systemDecimalSeparator = workbook.application.cultureInfo.numberFormat.numberDecimalSeparator;",
        ],
      },
      {
        name: "Excel.NumberFormatInfo.numberGroupSeparator",
        description:
          "Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on current system settings.",
        kind: "Property",
        signature: "Excel.NumberFormatInfo.numberGroupSeparator: string",
        examples: [
          "const systemThousandsSeparator = workbook.application.cultureInfo.numberFormat.numberGroupSeparator;",
        ],
      },
    ],
  },
  {
    objName: "Excel.NumberFormatProperty",
    apiList: [
      {
        name: "Excel.NumberFormatProperty.currency",
        description: "Indicates if the number format is of type currency.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.currency: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.dateTime",
        description: "Indicates if the number format is of type date-time.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.dateTime: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.dateTimeHasDayOfWeek",
        description: "Indicates if the date-time format has day-of-week.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.dateTimeHasDayOfWeek: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.dateTimeHasMonth",
        description: "Indicates if the date-time format has month.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.dateTimeHasMonth: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.dateTimeHasYear",
        description: "Indicates if the date-time format has year.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.dateTimeHasYear: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.key",
        description: "A key that corresponds to a number format.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.key: string",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.numeric",
        description: "Indicates if the number format is of type numeric.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.numeric: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.percent",
        description: "Indicates if the number format is of type percentage.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.percent: boolean",
        examples: [],
      },
      {
        name: "Excel.NumberFormatProperty.text",
        description: "Indicates if the number format is of type text.",
        kind: "Property",
        signature: "Excel.NumberFormatProperty.text: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NumberFormatPropertyCollection",
    apiList: [
      {
        name: "Excel.NumberFormatPropertyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.NumberFormatPropertyCollection.items: NumberFormatProperty[]",
        examples: [],
      },
      {
        name: "Excel.NumberFormatPropertyCollection.getItemAt",
        description: "Gets a `NumberFormatProperty` object by using its index in the collection.",
        kind: "Method",
        signature:
          "Excel.NumberFormatPropertyCollection.getItemAt => (index: number) => Excel.NumberFormatProperty",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.NumErrorCellValue",
    apiList: [
      {
        name: "Excel.NumErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.NumErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.NumErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.NumErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.NumErrorCellValue.errorSubType",
        description: "Represents the type of `NumErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.NumErrorCellValue.errorSubType: "Unknown" | NumErrorCellValueSubType | "ArrayTooLarge"',
        examples: [],
      },
      {
        name: "Excel.NumErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.NumErrorCellValue.errorType: ErrorCellValueType.num | "Num"',
        examples: [],
      },
      {
        name: "Excel.NumErrorCellValue.functionName",
        description: "Represents the name of the function causing the error.",
        kind: "Property",
        signature: "Excel.NumErrorCellValue.functionName: string",
        examples: [],
      },
      {
        name: "Excel.NumErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.NumErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PageBreak",
    apiList: [
      {
        name: "Excel.PageBreak.columnIndex",
        description: "Specifies the column index for the page break.",
        kind: "Property",
        signature: "Excel.PageBreak.columnIndex: number",
        examples: [],
      },
      {
        name: "Excel.PageBreak.rowIndex",
        description: "Specifies the row index for the page break.",
        kind: "Property",
        signature: "Excel.PageBreak.rowIndex: number",
        examples: [],
      },
      {
        name: "Excel.PageBreak.delete",
        description: "Deletes a page break object.",
        kind: "Method",
        signature: "Excel.PageBreak.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.PageBreak.getCellAfterBreak",
        description: "Gets the first cell after the page break.",
        kind: "Method",
        signature: "Excel.PageBreak.getCellAfterBreak => () => Excel.Range",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PageBreakCollection",
    apiList: [
      {
        name: "Excel.PageBreakCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PageBreakCollection.items: PageBreak[]",
        examples: [],
      },
      {
        name: "Excel.PageBreakCollection.add",
        description: "Adds a page break before the top-left cell of the range specified.",
        kind: "Method",
        signature:
          "Excel.PageBreakCollection.add(pageBreakRange: string | Excel.Range) => Excel.PageBreak",
        examples: ['activeWorksheet.horizontalPageBreaks.add("A21:E21");'],
      },
      {
        name: "Excel.PageBreakCollection.getCount",
        description: "Gets the number of page breaks in the collection.",
        kind: "Method",
        signature:
          "Excel.PageBreakCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PageBreakCollection.getItem",
        description: "Gets a page break object via the index.",
        kind: "Method",
        signature: "Excel.PageBreakCollection.getItem => (index: number) => Excel.PageBreak",
        examples: [],
      },
      {
        name: "Excel.PageBreakCollection.removePageBreaks",
        description: "Resets all manual page breaks in the collection.",
        kind: "Method",
        signature: "Excel.PageBreakCollection.removePageBreaks => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PageLayout",
    apiList: [
      {
        name: "Excel.PageLayout.blackAndWhite",
        description: "The worksheet's black and white print option.",
        kind: "Property",
        signature: "Excel.PageLayout.blackAndWhite: boolean",
        examples: [],
      },
      {
        name: "Excel.PageLayout.bottomMargin",
        description: "The worksheet's bottom page margin to use for printing in points.",
        kind: "Property",
        signature: "Excel.PageLayout.bottomMargin: number",
        examples: [],
      },
      {
        name: "Excel.PageLayout.centerHorizontally",
        description:
          "The worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.",
        kind: "Property",
        signature: "Excel.PageLayout.centerHorizontally: boolean",
        examples: ["activeWorksheet.pageLayout.centerHorizontally = true;"],
      },
      {
        name: "Excel.PageLayout.centerVertically",
        description:
          "The worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.",
        kind: "Property",
        signature: "Excel.PageLayout.centerVertically: boolean",
        examples: ["activeWorksheet.pageLayout.centerVertically = true;"],
      },
      {
        name: "Excel.PageLayout.draftMode",
        description:
          "The worksheet's draft mode option. If `true`, the sheet will be printed without graphics.",
        kind: "Property",
        signature: "Excel.PageLayout.draftMode: boolean",
        examples: [],
      },
      {
        name: "Excel.PageLayout.firstPageNumber",
        description:
          'The worksheet\'s first page number to print. A `null` value represents "auto" page numbering.',
        kind: "Property",
        signature: 'Excel.PageLayout.firstPageNumber: number | ""',
        examples: [],
      },
      {
        name: "Excel.PageLayout.footerMargin",
        description: "The worksheet's footer margin, in points, for use when printing.",
        kind: "Property",
        signature: "Excel.PageLayout.footerMargin: number",
        examples: [],
      },
      {
        name: "Excel.PageLayout.headerMargin",
        description: "The worksheet's header margin, in points, for use when printing.",
        kind: "Property",
        signature: "Excel.PageLayout.headerMargin: number",
        examples: [],
      },
      {
        name: "Excel.PageLayout.headersFooters",
        description: "Header and footer configuration for the worksheet.",
        kind: "Property",
        signature: "Excel.PageLayout.headersFooters: HeaderFooterGroup",
        examples: [],
      },
      {
        name: "Excel.PageLayout.leftMargin",
        description: "The worksheet's left margin, in points, for use when printing.",
        kind: "Property",
        signature: "Excel.PageLayout.leftMargin: number",
        examples: [],
      },
      {
        name: "Excel.PageLayout.orientation",
        description: "The worksheet's orientation of the page.",
        kind: "Property",
        signature: 'Excel.PageLayout.orientation: Excel.PageOrientation | "Portrait" | "Landscape"',
        examples: ["activeWorksheet.pageLayout.orientation = Excel.PageOrientation.landscape;"],
      },
      {
        name: "Excel.PageLayout.paperSize",
        description: "The worksheet's paper size of the page.",
        kind: "Property",
        signature:
          'Excel.PageLayout.paperSize: PaperType | "Letter" | "LetterSmall" | "Tabloid" | "Ledger" | "Legal" | "Statement" | "Executive" | "A3" | "A4" | "A4Small" | "A5" | "B4" | "B5" | "Folio" | "Quatro" | ... 25 more ... | "FanfoldLegalGerman"',
        examples: [],
      },
      {
        name: "Excel.PageLayout.printComments",
        description: "Specifies if the worksheet's comments should be displayed when printing.",
        kind: "Property",
        signature:
          'Excel.PageLayout.printComments: PrintComments | "NoComments" | "EndSheet" | "InPlace"',
        examples: [],
      },
      {
        name: "Excel.PageLayout.printErrors",
        description: "The worksheet's print errors option.",
        kind: "Property",
        signature:
          'Excel.PageLayout.printErrors: "NotAvailable" | "Dash" | PrintErrorType | "AsDisplayed" | "Blank"',
        examples: [],
      },
      {
        name: "Excel.PageLayout.printGridlines",
        description: "Specifies if the worksheet's gridlines will be printed.",
        kind: "Property",
        signature: "Excel.PageLayout.printGridlines: boolean",
        examples: [],
      },
      {
        name: "Excel.PageLayout.printHeadings",
        description: "Specifies if the worksheet's headings will be printed.",
        kind: "Property",
        signature: "Excel.PageLayout.printHeadings: boolean",
        examples: [],
      },
      {
        name: "Excel.PageLayout.printOrder",
        description:
          "The worksheet's page print order option. This specifies the order to use for processing the page number printed.",
        kind: "Property",
        signature: 'Excel.PageLayout.printOrder: PrintOrder | "DownThenOver" | "OverThenDown"',
        examples: [],
      },
      {
        name: "Excel.PageLayout.rightMargin",
        description: "The worksheet's right margin, in points, for use when printing.",
        kind: "Property",
        signature: "Excel.PageLayout.rightMargin: number",
        examples: [],
      },
      {
        name: "Excel.PageLayout.topMargin",
        description: "The worksheet's top margin, in points, for use when printing.",
        kind: "Property",
        signature: "Excel.PageLayout.topMargin: number",
        examples: [],
      },
      {
        name: "Excel.PageLayout.zoom",
        description:
          "The worksheet's print zoom options. The `PageLayoutZoomOptions` object must be set as a JSON object (use `x.zoom = {...}` instead of `x.zoom.scale = ...`).",
        kind: "Property",
        signature: "Excel.PageLayout.zoom: Excel.PageLayoutZoomOptions",
        examples: ["activeWorksheet.pageLayout.zoom = { scale: 200 };"],
      },
      {
        name: "Excel.PageLayout.getPrintArea",
        description:
          "Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, an `ItemNotFound` error will be thrown.",
        kind: "Method",
        signature: "Excel.PageLayout.getPrintArea => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.PageLayout.getPrintAreaOrNullObject",
        description:
          "Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.PageLayout.getPrintAreaOrNullObject => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.PageLayout.getPrintTitleColumns",
        description: "Gets the range object representing the title columns.",
        kind: "Method",
        signature: "Excel.PageLayout.getPrintTitleColumns => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PageLayout.getPrintTitleColumnsOrNullObject",
        description:
          "Gets the range object representing the title columns. If not set, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.PageLayout.getPrintTitleColumnsOrNullObject => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PageLayout.getPrintTitleRows",
        description: "Gets the range object representing the title rows.",
        kind: "Method",
        signature: "Excel.PageLayout.getPrintTitleRows => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PageLayout.getPrintTitleRowsOrNullObject",
        description:
          "Gets the range object representing the title rows. If not set, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.PageLayout.getPrintTitleRowsOrNullObject => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PageLayout.setPrintArea",
        description: "Sets the worksheet's print area.",
        kind: "Method",
        signature:
          "Excel.PageLayout.setPrintArea(printArea: string | Excel.Range | Excel.RangeAreas) => void",
        examples: [
          'activeWorksheet.pageLayout.setPrintArea("A1:D100");',
          'activeWorksheet.pageLayout.setPrintArea("A1:D41");',
        ],
      },
      {
        name: "Excel.PageLayout.setPrintMargins",
        description: "Sets the worksheet's page margins with units.",
        kind: "Method",
        signature:
          'Excel.PageLayout.setPrintMargins => { (unit: PrintMarginUnit, marginOptions: PageLayoutMarginOptions): void; (unit: "Points" | "Inches" | "Centimeters", marginOptions: PageLayoutMarginOptions): void; (unit: string, marginOptions: Excel.PageLayoutMarginOptions): void; }',
        examples: [],
      },
      {
        name: "Excel.PageLayout.setPrintTitleColumns",
        description:
          "Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.",
        kind: "Method",
        signature:
          "Excel.PageLayout.setPrintTitleColumns => (printTitleColumns: Range | string) => void",
        examples: [],
      },
      {
        name: "Excel.PageLayout.setPrintTitleRows",
        description:
          "Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.",
        kind: "Method",
        signature:
          "Excel.PageLayout.setPrintTitleRows(printTitleRows: string | Excel.Range) => void",
        examples: ['activeWorksheet.pageLayout.setPrintTitleRows("$1:$1");'],
      },
    ],
  },
  {
    objName: "Excel.PageLayoutMarginOptions",
    apiList: [
      {
        name: "Excel.PageLayoutMarginOptions.bottom",
        description:
          "Specifies the page layout bottom margin in the unit specified to use for printing.",
        kind: "Property",
        signature: "Excel.PageLayoutMarginOptions.bottom: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutMarginOptions.footer",
        description:
          "Specifies the page layout footer margin in the unit specified to use for printing.",
        kind: "Property",
        signature: "Excel.PageLayoutMarginOptions.footer: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutMarginOptions.header",
        description:
          "Specifies the page layout header margin in the unit specified to use for printing.",
        kind: "Property",
        signature: "Excel.PageLayoutMarginOptions.header: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutMarginOptions.left",
        description:
          "Specifies the page layout left margin in the unit specified to use for printing.",
        kind: "Property",
        signature: "Excel.PageLayoutMarginOptions.left: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutMarginOptions.right",
        description:
          "Specifies the page layout right margin in the unit specified to use for printing.",
        kind: "Property",
        signature: "Excel.PageLayoutMarginOptions.right: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutMarginOptions.top",
        description:
          "Specifies the page layout top margin in the unit specified to use for printing.",
        kind: "Property",
        signature: "Excel.PageLayoutMarginOptions.top: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PageLayoutZoomOptions",
    apiList: [
      {
        name: "Excel.PageLayoutZoomOptions.horizontalFitToPages",
        description:
          "Number of pages to fit horizontally. This value can be `null` if percentage scale is used.",
        kind: "Property",
        signature: "Excel.PageLayoutZoomOptions.horizontalFitToPages: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutZoomOptions.scale",
        description:
          "Print page scale value can be between 10 and 400. This value can be `null` if fit to page tall or wide is specified.",
        kind: "Property",
        signature: "Excel.PageLayoutZoomOptions.scale: number",
        examples: [],
      },
      {
        name: "Excel.PageLayoutZoomOptions.verticalFitToPages",
        description:
          "Number of pages to fit vertically. This value can be `null` if percentage scale is used.",
        kind: "Property",
        signature: "Excel.PageLayoutZoomOptions.verticalFitToPages: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotDateFilter",
    apiList: [
      {
        name: "Excel.PivotDateFilter.comparator",
        description:
          "The comparator is the static value to which other values are compared. The type of comparison is defined by the condition.",
        kind: "Property",
        signature: "Excel.PivotDateFilter.comparator: FilterDatetime",
        examples: [],
      },
      {
        name: "Excel.PivotDateFilter.condition",
        description:
          "Specifies the condition for the filter, which defines the necessary filtering criteria.",
        kind: "Property",
        signature:
          'Excel.PivotDateFilter.condition: "Unknown" | DateFilterCondition | "Equals" | "Before" | "BeforeOrEqualTo" | "After" | "AfterOrEqualTo" | "Between" | "Tomorrow" | "Today" | "Yesterday" | ... 28 more ... | "AllDatesInPeriodDecember"',
        examples: [],
      },
      {
        name: "Excel.PivotDateFilter.exclusive",
        description:
          "If `true`, filter *excludes* items that meet criteria. The default is `false` (filter to include items that meet criteria).",
        kind: "Property",
        signature: "Excel.PivotDateFilter.exclusive: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotDateFilter.lowerBound",
        description: "The lower-bound of the range for the `between` filter condition.",
        kind: "Property",
        signature: "Excel.PivotDateFilter.lowerBound: FilterDatetime",
        examples: [],
      },
      {
        name: "Excel.PivotDateFilter.upperBound",
        description: "The upper-bound of the range for the `between` filter condition.",
        kind: "Property",
        signature: "Excel.PivotDateFilter.upperBound: FilterDatetime",
        examples: [],
      },
      {
        name: "Excel.PivotDateFilter.wholeDays",
        description:
          "For `equals`, `before`, `after`, and `between` filter conditions, indicates if comparisons should be made as whole days.",
        kind: "Property",
        signature: "Excel.PivotDateFilter.wholeDays: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotField",
    apiList: [
      {
        name: "Excel.PivotField.id",
        description: "ID of the PivotField.",
        kind: "Property",
        signature: "Excel.PivotField.id: string",
        examples: [],
      },
      {
        name: "Excel.PivotField.items",
        description: "Returns the PivotItems associated with the PivotField.",
        kind: "Property",
        signature: "Excel.PivotField.items: Excel.PivotItemCollection",
        examples: [
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
      {
        name: "Excel.PivotField.name",
        description: "Name of the PivotField.",
        kind: "Property",
        signature: "Excel.PivotField.name: string",
        examples: [],
      },
      {
        name: "Excel.PivotField.showAllItems",
        description: "Determines whether to show all items of the PivotField.",
        kind: "Property",
        signature: "Excel.PivotField.showAllItems: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotField.subtotals",
        description: "Subtotals of the PivotField.",
        kind: "Property",
        signature: "Excel.PivotField.subtotals: Subtotals",
        examples: [],
      },
      {
        name: "Excel.PivotField.applyFilter",
        description:
          "Sets one or more of the field's current PivotFilters and applies them to the field. If the provided filters are invalid or cannot be applied, an exception is thrown.",
        kind: "Method",
        signature: "Excel.PivotField.applyFilter(filter: Excel.PivotFilters) => void",
        examples: [
          "filterField.applyFilter({ dateFilter: dateFilter });",
          "field.applyFilter({ labelFilter: filter });",
          "filterField.applyFilter({ manualFilter: manualFilter });",
          "field.applyFilter({ valueFilter: filter });",
        ],
      },
      {
        name: "Excel.PivotField.clearAllFilters",
        description:
          "Clears all criteria from all of the field's filters. This removes any active filtering on the field.",
        kind: "Method",
        signature: "Excel.PivotField.clearAllFilters() => void",
        examples: ["hierarchy.fields.getItem(hierarchy.name).clearAllFilters();"],
      },
      {
        name: "Excel.PivotField.clearFilter",
        description:
          "Clears all existing criteria from the field's filter of the given type (if one is currently applied).",
        kind: "Method",
        signature:
          'Excel.PivotField.clearFilter => { (filterType: PivotFilterType): void; (filterType: "Unknown" | "Value" | "Manual" | "Date" | "Label"): void; (filterType: string): void; }',
        examples: [],
      },
      {
        name: "Excel.PivotField.getFilters",
        description: "Gets all filters currently applied on the field.",
        kind: "Method",
        signature:
          "Excel.PivotField.getFilters => () => OfficeExtension.ClientResult<Excel.PivotFilters>",
        examples: [],
      },
      {
        name: "Excel.PivotField.isFiltered",
        description: "Checks if there are any applied filters on the field.",
        kind: "Method",
        signature:
          'Excel.PivotField.isFiltered => { (filterType?: PivotFilterType): OfficeExtension.ClientResult<boolean>; (filterType?: "Unknown" | "Value" | "Manual" | "Date" | "Label"): OfficeExtension.ClientResult<...>; (filterType?: string): OfficeExtension.ClientResult<boolean>; }',
        examples: [],
      },
      {
        name: "Excel.PivotField.sortByLabels",
        description:
          "Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.",
        kind: "Method",
        signature: "Excel.PivotField.sortByLabels => (sortBy: SortBy) => void",
        examples: [],
      },
      {
        name: "Excel.PivotField.sortByValues",
        description:
          "Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when there are multiple values from the same DataPivotHierarchy.",
        kind: "Method",
        signature:
          'Excel.PivotField.sortByValues => { (sortBy: SortBy, valuesHierarchy: DataPivotHierarchy, pivotItemScope?: (string | PivotItem)[]): void; (sortBy: "Ascending" | "Descending", valuesHierarchy: DataPivotHierarchy, pivotItemScope?: (string | PivotItem)[]): void; (sortBy: string, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotIte...',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotFieldCollection",
    apiList: [
      {
        name: "Excel.PivotFieldCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PivotFieldCollection.items: PivotField[]",
        examples: [],
      },
      {
        name: "Excel.PivotFieldCollection.getCount",
        description: "Gets the number of pivot fields in the collection.",
        kind: "Method",
        signature:
          "Excel.PivotFieldCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PivotFieldCollection.getItem",
        description: "Gets a PivotField by its name or ID.",
        kind: "Method",
        signature: "Excel.PivotFieldCollection.getItem(name: string) => Excel.PivotField",
        examples: [
          'let filterField = dateHierarchy.fields.getItem("Date Updated");',
          "hierarchy.fields.getItem(hierarchy.name).clearAllFilters();",
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'const filterField = dateHierarchy.fields.getItem("Date Updated");',
          'const field = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");',
          'const filterField = classHierarchy.fields.getItem("Classification");',
          'const field = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
    ],
  },
  {
    objName: "Excel.PivotFilters",
    apiList: [
      {
        name: "Excel.PivotFilters.dateFilter",
        description:
          "The PivotField's currently applied date filter. This property is `null` if no value filter is applied.",
        kind: "Property",
        signature: "Excel.PivotFilters.dateFilter: PivotDateFilter",
        examples: [],
      },
      {
        name: "Excel.PivotFilters.labelFilter",
        description:
          "The PivotField's currently applied label filter. This property is `null` if no value filter is applied.",
        kind: "Property",
        signature: "Excel.PivotFilters.labelFilter: PivotLabelFilter",
        examples: [],
      },
      {
        name: "Excel.PivotFilters.manualFilter",
        description:
          "The PivotField's currently applied manual filter. This property is `null` if no value filter is applied.",
        kind: "Property",
        signature: "Excel.PivotFilters.manualFilter: PivotManualFilter",
        examples: [],
      },
      {
        name: "Excel.PivotFilters.valueFilter",
        description:
          "The PivotField's currently applied value filter. This property is `null` if no value filter is applied.",
        kind: "Property",
        signature: "Excel.PivotFilters.valueFilter: PivotValueFilter",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotHierarchy",
    apiList: [
      {
        name: "Excel.PivotHierarchy.fields",
        description: "Returns the PivotFields associated with the PivotHierarchy.",
        kind: "Property",
        signature: "Excel.PivotHierarchy.fields: Excel.PivotFieldCollection",
        examples: [
          "hierarchy.fields.getItem(hierarchy.name).clearAllFilters();",
          'const field = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");',
          'const field = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");',
        ],
      },
      {
        name: "Excel.PivotHierarchy.id",
        description: "ID of the PivotHierarchy.",
        kind: "Property",
        signature: "Excel.PivotHierarchy.id: string",
        examples: [],
      },
      {
        name: "Excel.PivotHierarchy.name",
        description: "Name of the PivotHierarchy.",
        kind: "Property",
        signature: "Excel.PivotHierarchy.name: string",
        examples: ["hierarchy.fields.getItem(hierarchy.name).clearAllFilters();"],
      },
    ],
  },
  {
    objName: "Excel.PivotHierarchyCollection",
    apiList: [
      {
        name: "Excel.PivotHierarchyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PivotHierarchyCollection.items: Excel.PivotHierarchy[]",
        examples: [
          "pivotTable.hierarchies.items.forEach(function (hierarchy) {\n    hierarchy.fields.getItem(hierarchy.name).clearAllFilters();\n  });",
          "pivotTable.hierarchies.items.forEach((hierarchy) => {\n    hierarchy.fields.getItem(hierarchy.name).clearAllFilters();\n  });",
        ],
      },
      {
        name: "Excel.PivotHierarchyCollection.getCount",
        description: "Gets the number of pivot hierarchies in the collection.",
        kind: "Method",
        signature:
          "Excel.PivotHierarchyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PivotHierarchyCollection.getItem",
        description: "Gets a PivotHierarchy by its name or ID.",
        kind: "Method",
        signature: "Excel.PivotHierarchyCollection.getItem(name: string) => Excel.PivotHierarchy",
        examples: [
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
          'pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));',
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));',
          'dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));',
          'const field = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");',
          'classHierarchy = pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
          'const field = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");',
        ],
      },
    ],
  },
  {
    objName: "Excel.PivotItem",
    apiList: [
      {
        name: "Excel.PivotItem.id",
        description: "ID of the PivotItem.",
        kind: "Property",
        signature: "Excel.PivotItem.id: string",
        examples: [],
      },
      {
        name: "Excel.PivotItem.isExpanded",
        description:
          "Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.",
        kind: "Property",
        signature: "Excel.PivotItem.isExpanded: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotItem.name",
        description: "Name of the PivotItem.",
        kind: "Property",
        signature: "Excel.PivotItem.name: string",
        examples: [],
      },
      {
        name: "Excel.PivotItem.visible",
        description: "Specifies if the PivotItem is visible.",
        kind: "Property",
        signature: "Excel.PivotItem.visible: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotItemCollection",
    apiList: [
      {
        name: "Excel.PivotItemCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PivotItemCollection.items: PivotItem[]",
        examples: [],
      },
      {
        name: "Excel.PivotItemCollection.getCount",
        description: "Gets the number of PivotItems in the collection.",
        kind: "Method",
        signature:
          "Excel.PivotItemCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PivotItemCollection.getItem",
        description: "Gets a PivotItem by its name or ID.",
        kind: "Method",
        signature: "Excel.PivotItemCollection.getItem(name: string) => Excel.PivotItem",
        examples: [
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
    ],
  },
  {
    objName: "Excel.PivotLabelFilter",
    apiList: [
      {
        name: "Excel.PivotLabelFilter.comparator",
        description:
          "The comparator is the static value to which other values are compared. The type of comparison is defined by the condition. Note: A numeric string is treated as a number when being compared against other numeric strings.",
        kind: "Property",
        signature: "Excel.PivotLabelFilter.comparator: string",
        examples: [],
      },
      {
        name: "Excel.PivotLabelFilter.condition",
        description:
          "Specifies the condition for the filter, which defines the necessary filtering criteria.",
        kind: "Property",
        signature:
          'Excel.PivotLabelFilter.condition: "Unknown" | LabelFilterCondition | "Equals" | "Between" | "BeginsWith" | "EndsWith" | "Contains" | "GreaterThan" | "GreaterThanOrEqualTo" | "LessThan" | "LessThanOrEqualTo"',
        examples: [],
      },
      {
        name: "Excel.PivotLabelFilter.exclusive",
        description:
          "If `true`, filter *excludes* items that meet criteria. The default is `false` (filter to include items that meet criteria).",
        kind: "Property",
        signature: "Excel.PivotLabelFilter.exclusive: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotLabelFilter.lowerBound",
        description:
          "The lower-bound of the range for the `between` filter condition. Note: A numeric string is treated as a number when being compared against other numeric strings.",
        kind: "Property",
        signature: "Excel.PivotLabelFilter.lowerBound: string",
        examples: [],
      },
      {
        name: "Excel.PivotLabelFilter.substring",
        description:
          "The substring used for `beginsWith`, `endsWith`, and `contains` filter conditions.",
        kind: "Property",
        signature: "Excel.PivotLabelFilter.substring: string",
        examples: [],
      },
      {
        name: "Excel.PivotLabelFilter.upperBound",
        description:
          "The upper-bound of the range for the `between` filter condition. Note: A numeric string is treated as a number when being compared against other numeric strings.",
        kind: "Property",
        signature: "Excel.PivotLabelFilter.upperBound: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotLayout",
    apiList: [
      {
        name: "Excel.PivotLayout.altTextDescription",
        description:
          "The alt text description of the PivotTable. Alt text provides alternative, text-based representations of the information contained in the PivotTable. This information is useful for people with vision or cognitive impairments who may not be able to see or understand the table. A title can be read to a person with a disability and is used to determine whether they wish to hear the description of the content.",
        kind: "Property",
        signature: "Excel.PivotLayout.altTextDescription: string",
        examples: [
          'pivotLayout.altTextDescription =\n    "A summary of fruit sales. It is pivoted on farm name, and fruit type. The aggregated data is both the sums of crates sold at the farms and the sums of crates sold wholesale.";',
        ],
      },
      {
        name: "Excel.PivotLayout.altTextTitle",
        description:
          "The alt text title of the PivotTable. Alt text provides alternative, text-based representations of the information contained in the PivotTable. This information is useful for people with vision or cognitive impairments who may not be able to see or understand the table. A title can be read to a person with a disability and is used to determine whether they wish to hear the description of the content.",
        kind: "Property",
        signature: "Excel.PivotLayout.altTextTitle: string",
        examples: ['pivotLayout.altTextTitle = "Farm Sales PivotTable";'],
      },
      {
        name: "Excel.PivotLayout.autoFormat",
        description:
          "Specifies if formatting will be automatically formatted when itâ€™s refreshed or when fields are moved.",
        kind: "Property",
        signature: "Excel.PivotLayout.autoFormat: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.emptyCellText",
        description:
          "The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`. Note that this value persists if `fillEmptyCells` is set to `false`, and that setting this value does not set that property to `true`. By default, this is an empty string.",
        kind: "Property",
        signature: "Excel.PivotLayout.emptyCellText: string",
        examples: ['pivotLayout.emptyCellText = "--";'],
      },
      {
        name: "Excel.PivotLayout.enableFieldList",
        description: "Specifies if the field list can be shown in the UI.",
        kind: "Property",
        signature: "Excel.PivotLayout.enableFieldList: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.fillEmptyCells",
        description:
          "Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`. Default is `false`. Note that the value of `emptyCellText` persists when this property is set to `false`.",
        kind: "Property",
        signature: "Excel.PivotLayout.fillEmptyCells: boolean",
        examples: [
          "pivotLayout.fillEmptyCells = true;",
          "let fillToSet = !pivotLayout.fillEmptyCells;",
          "pivotLayout.fillEmptyCells = fillToSet;",
        ],
      },
      {
        name: "Excel.PivotLayout.layoutType",
        description:
          "This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.",
        kind: "Property",
        signature:
          'Excel.PivotLayout.layoutType: Excel.PivotLayoutType | "Compact" | "Tabular" | "Outline"',
        examples: [
          'pivotTable.layout.layoutType = "Outline";',
          'pivotTable.layout.layoutType = "Tabular";',
          'pivotTable.layout.layoutType = "Compact";',
          '"Pivot layout is now " + pivotTable.layout.layoutType;',
        ],
      },
      {
        name: "Excel.PivotLayout.pivotStyle",
        description: "The style applied to the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotLayout.pivotStyle: PivotTableStyle",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.preserveFormatting",
        description:
          "Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.",
        kind: "Property",
        signature: "Excel.PivotLayout.preserveFormatting: boolean",
        examples: [
          "pivotLayout.preserveFormatting = true;",
          "let preserveFormattingToSet = !pivotLayout.preserveFormatting;",
          "pivotLayout.preserveFormatting = preserveFormattingToSet;",
        ],
      },
      {
        name: "Excel.PivotLayout.showColumnGrandTotals",
        description: "Specifies if the PivotTable report shows grand totals for columns.",
        kind: "Property",
        signature: "Excel.PivotLayout.showColumnGrandTotals: boolean",
        examples: [
          "let showColumnTotals = !pivotLayout.showColumnGrandTotals;",
          "pivotLayout.showColumnGrandTotals = showColumnTotals;",
        ],
      },
      {
        name: "Excel.PivotLayout.showFieldHeaders",
        description:
          "Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).",
        kind: "Property",
        signature: "Excel.PivotLayout.showFieldHeaders: boolean",
        examples: [
          "let showHeaders = !pivotLayout.showFieldHeaders;",
          "pivotLayout.showFieldHeaders = showHeaders;",
        ],
      },
      {
        name: "Excel.PivotLayout.showRowGrandTotals",
        description: "Specifies if the PivotTable report shows grand totals for rows.",
        kind: "Property",
        signature: "Excel.PivotLayout.showRowGrandTotals: boolean",
        examples: [
          "let showRowTotals = !pivotLayout.showRowGrandTotals;",
          "pivotLayout.showRowGrandTotals = showRowTotals;",
        ],
      },
      {
        name: "Excel.PivotLayout.subtotalLocation",
        description:
          "This property indicates the `SubtotalLocationType` of all fields on the PivotTable. If fields have different states, this will be `null`.",
        kind: "Property",
        signature:
          'Excel.PivotLayout.subtotalLocation: SubtotalLocationType | "AtTop" | "AtBottom" | "Off"',
        examples: [],
      },
      {
        name: "Excel.PivotLayout.tabularNumberFormat",
        description:
          "Returns a 2D array that contains pivot table's cell number format strings in tabular layout and no sub/grand totals.",
        kind: "Property",
        signature: "Excel.PivotLayout.tabularNumberFormat: any[][]",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.tabularNumberFormatLocal",
        description:
          "Returns a 2D array that contains pivot table's cell local number format strings in tabular layout and no sub/grand totals.",
        kind: "Property",
        signature: "Excel.PivotLayout.tabularNumberFormatLocal: any[][]",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.tabularText",
        description:
          "Returns a 2D array that contains pivot table's cell display texts in tabular layout and no sub/grand totals.",
        kind: "Property",
        signature: "Excel.PivotLayout.tabularText: any[][]",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.tabularValues",
        description:
          "Returns a 2D array that contains pivot table's cell values in tabular layout and no sub/grand totals.",
        kind: "Property",
        signature: "Excel.PivotLayout.tabularValues: any[][]",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.displayBlankLineAfterEachItem",
        description:
          "Sets whether or not to display a blank line after each item. This is set at the global level for the PivotTable and applied to individual PivotFields. This function overwrites the setting for all fields in the PivotTable to the value of `display` parameter.",
        kind: "Method",
        signature: "Excel.PivotLayout.displayBlankLineAfterEachItem(display: boolean) => void",
        examples: ["pivotLayout.displayBlankLineAfterEachItem(true);"],
      },
      {
        name: "Excel.PivotLayout.getCell",
        description:
          "Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies. The returned cell is the intersection of the given row and column that contains the data from the given hierarchy. This method is the inverse of calling `getPivotItems` and `getDataHierarchy` on a particular cell.",
        kind: "Method",
        signature:
          "Excel.PivotLayout.getCell => (dataHierarchy: DataPivotHierarchy | string, rowItems: Array<PivotItem | string>, columnItems: Array<PivotItem | string>) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.getColumnLabelRange",
        description: "Returns the range where the PivotTable's column labels reside.",
        kind: "Method",
        signature: "Excel.PivotLayout.getColumnLabelRange => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.getDataBodyRange",
        description: "Returns the range where the PivotTable's data values reside.",
        kind: "Method",
        signature: "Excel.PivotLayout.getDataBodyRange() => Excel.Range",
        examples: [
          "let range = pivotTable.layout.getDataBodyRange();",
          "pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;",
          "const range = pivotTable.layout.getDataBodyRange();",
        ],
      },
      {
        name: "Excel.PivotLayout.getDataHierarchy",
        description:
          "Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.",
        kind: "Method",
        signature:
          "Excel.PivotLayout.getDataHierarchy => (cell: Range | string) => Excel.DataPivotHierarchy",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.getFilterAxisRange",
        description: "Returns the range of the PivotTable's filter area.",
        kind: "Method",
        signature: "Excel.PivotLayout.getFilterAxisRange => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.getPivotItems",
        description:
          "Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.",
        kind: "Method",
        signature:
          'Excel.PivotLayout.getPivotItems => { (axis: PivotAxis, cell: string | Range): PivotItemCollection; (axis: "Unknown" | "Column" | "Row" | "Data" | "Filter", cell: string | Range): PivotItemCollection; (axis: string, cell: Range | string): Excel.PivotItemCollection; }',
        examples: [],
      },
      {
        name: "Excel.PivotLayout.getRange",
        description: "Returns the range the PivotTable exists on, excluding the filter area.",
        kind: "Method",
        signature: "Excel.PivotLayout.getRange => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.getRowLabelRange",
        description: "Returns the range where the PivotTable's row labels reside.",
        kind: "Method",
        signature: "Excel.PivotLayout.getRowLabelRange => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.PivotLayout.repeatAllItemLabels",
        description:
          'Sets the "repeat all item labels" setting across all fields in the PivotTable.',
        kind: "Method",
        signature: "Excel.PivotLayout.repeatAllItemLabels(repeatLabels: boolean) => void",
        examples: ["pivotLayout.repeatAllItemLabels(true);"],
      },
      {
        name: "Excel.PivotLayout.setAutoSortOnCell",
        description:
          "Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.",
        kind: "Method",
        signature:
          'Excel.PivotLayout.setAutoSortOnCell => { (cell: string | Range, sortBy: SortBy): void; (cell: string | Range, sortBy: "Ascending" | "Descending"): void; (cell: Range | string, sortBy: string): void; }',
        examples: [],
      },
      {
        name: "Excel.PivotLayout.setStyle",
        description: "Sets the style applied to the PivotTable.",
        kind: "Method",
        signature:
          "Excel.PivotLayout.setStyle => (style: string | PivotTableStyle | BuiltInPivotTableStyle) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotManualFilter",
    apiList: [
      {
        name: "Excel.PivotManualFilter.selectedItems",
        description:
          "A list of selected items to manually filter. These must be existing and valid items from the chosen field.",
        kind: "Property",
        signature: "Excel.PivotManualFilter.selectedItems: (string | PivotItem)[]",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotTable",
    apiList: [
      {
        name: "Excel.PivotTable.allowMultipleFiltersPerField",
        description:
          "Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.",
        kind: "Property",
        signature: "Excel.PivotTable.allowMultipleFiltersPerField: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotTable.columnHierarchies",
        description: "The Column Pivot Hierarchies of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.columnHierarchies: Excel.RowColumnPivotHierarchyCollection",
        examples: [
          'pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'const column = pivotTable.columnHierarchies.getItemOrNullObject("Farm");',
          "pivotTable.columnHierarchies.remove(column);",
        ],
      },
      {
        name: "Excel.PivotTable.dataHierarchies",
        description: "The Data Pivot Hierarchies of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.dataHierarchies: Excel.DataPivotHierarchyCollection",
        examples: [
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));',
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));',
          "pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;",
          "pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;",
          'let farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");',
          "let dataHierarchies = pivotTable.dataHierarchies;",
          "const dataHierarchies = pivotTable.dataHierarchies;",
          'const farmDataHierarchy: Excel.DataPivotHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");',
        ],
      },
      {
        name: "Excel.PivotTable.enableDataValueEditing",
        description:
          "Specifies if the PivotTable allows values in the data body to be edited by the user.",
        kind: "Property",
        signature: "Excel.PivotTable.enableDataValueEditing: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotTable.filterHierarchies",
        description: "The Filter Pivot Hierarchies of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.filterHierarchies: Excel.FilterPivotHierarchyCollection",
        examples: [
          'let classHierarchy = pivotTable.filterHierarchies.getItemOrNullObject("Classification");',
          'classHierarchy = pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
        ],
      },
      {
        name: "Excel.PivotTable.hierarchies",
        description: "The Pivot Hierarchies of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.hierarchies: Excel.PivotHierarchyCollection",
        examples: [
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
          'pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));',
          'pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));',
          'dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));',
          'const field = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");',
          'classHierarchy = pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
          'const field = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");',
        ],
      },
      {
        name: "Excel.PivotTable.id",
        description: "ID of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.id: string",
        examples: [],
      },
      {
        name: "Excel.PivotTable.layout",
        description:
          "The PivotLayout describing the layout and visual structure of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.layout: Excel.PivotLayout",
        examples: [
          "let range = pivotTable.layout.getDataBodyRange();",
          'pivotTable.layout.layoutType = "Outline";',
          'pivotTable.layout.layoutType = "Tabular";',
          'pivotTable.layout.layoutType = "Compact";',
          "let pivotLayout = pivotTable.layout;",
          "const pivotLayout = pivotTable.layout;",
          "const range = pivotTable.layout.getDataBodyRange();",
          '"Pivot layout is now " + pivotTable.layout.layoutType;',
        ],
      },
      {
        name: "Excel.PivotTable.name",
        description: "Name of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.name: string",
        examples: [],
      },
      {
        name: "Excel.PivotTable.refreshOnOpen",
        description:
          'Specifies whether the PivotTable refreshes when the workbook opens. Corresponds to "Refresh on load" setting in the UI.',
        kind: "Property",
        signature: "Excel.PivotTable.refreshOnOpen: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotTable.rowHierarchies",
        description: "The Row Pivot Hierarchies of the PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.rowHierarchies: Excel.RowColumnPivotHierarchyCollection",
        examples: [
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'let dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");',
          'dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
      {
        name: "Excel.PivotTable.useCustomSortLists",
        description: "Specifies if the PivotTable uses custom lists when sorting.",
        kind: "Property",
        signature: "Excel.PivotTable.useCustomSortLists: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotTable.worksheet",
        description: "The worksheet containing the current PivotTable.",
        kind: "Property",
        signature: "Excel.PivotTable.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.PivotTable.addDateGroup",
        description: "Add grouping based on a DateTime Pivot Field.",
        kind: "Method",
        signature:
          'Excel.PivotTable.addDateGroup => { (pivotField: PivotField, groupBy: PivotTableDateGroupBy): PivotHierarchy; (pivotField: PivotField, groupBy: "Invalid" | ... 6 more ... | "ByYears"): PivotHierarchy; (pivotField: Excel.PivotField, groupBy: string): Excel.PivotHierarchy; }',
        examples: [],
      },
      {
        name: "Excel.PivotTable.delete",
        description: "Deletes the PivotTable.",
        kind: "Method",
        signature: "Excel.PivotTable.delete() => void",
        examples: ["pivotTable.delete();"],
      },
      {
        name: "Excel.PivotTable.getDataSourceString",
        description:
          "Returns the string representation of the data source for the PivotTable. This method currently supports string representations for table and range objects. Otherwise, it returns an empty string.",
        kind: "Method",
        signature: "Excel.PivotTable.getDataSourceString() => OfficeExtension.ClientResult<string>",
        examples: ["const pivotTableDataSourceString = pivotTable.getDataSourceString();"],
      },
      {
        name: "Excel.PivotTable.getDataSourceType",
        description: "Gets the type of the data source for the PivotTable.",
        kind: "Method",
        signature:
          "Excel.PivotTable.getDataSourceType() => OfficeExtension.ClientResult<Excel.DataSourceType>",
        examples: ["const pivotTableDataSourceType = pivotTable.getDataSourceType();"],
      },
      {
        name: "Excel.PivotTable.refresh",
        description: "Refreshes the PivotTable.",
        kind: "Method",
        signature: "Excel.PivotTable.refresh() => void",
        examples: ["pivotTable.refresh();"],
      },
    ],
  },
  {
    objName: "Excel.PivotTableCollection",
    apiList: [
      {
        name: "Excel.PivotTableCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PivotTableCollection.items: PivotTable[]",
        examples: [],
      },
      {
        name: "Excel.PivotTableCollection.add",
        description:
          "Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.",
        kind: "Method",
        signature:
          "Excel.PivotTableCollection.add(name: string, source: string | Excel.Range | Excel.Table, destination: string | Excel.Range) => Excel.PivotTable",
        examples: [
          'activeWorksheet.pivotTables.add("Farm Sales", "A1:E21", "A22");',
          'workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
          'workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");',
          'workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
        ],
      },
      {
        name: "Excel.PivotTableCollection.getCount",
        description: "Gets the number of pivot tables in the collection.",
        kind: "Method",
        signature:
          "Excel.PivotTableCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PivotTableCollection.getItem",
        description: "Gets a PivotTable by name.",
        kind: "Method",
        signature: "Excel.PivotTableCollection.getItem(name: string) => Excel.PivotTable",
        examples: [
          'const pivotTable = activeWorksheet.pivotTables.getItem("Farm Sales");',
          'const pivotTable = activeWorksheet.pivotTables.getItem("All Farm Sales");',
        ],
      },
      {
        name: "Excel.PivotTableCollection.refreshAll",
        description: "Refreshes all the pivot tables in the collection.",
        kind: "Method",
        signature: "Excel.PivotTableCollection.refreshAll => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotTableScopedCollection",
    apiList: [
      {
        name: "Excel.PivotTableScopedCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PivotTableScopedCollection.items: PivotTable[]",
        examples: [],
      },
      {
        name: "Excel.PivotTableScopedCollection.getCount",
        description: "Gets the number of PivotTables in the collection.",
        kind: "Method",
        signature:
          "Excel.PivotTableScopedCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PivotTableScopedCollection.getFirst",
        description:
          "Gets the first PivotTable in the collection. The PivotTables in the collection are sorted top-to-bottom and left-to-right, such that top-left table is the first PivotTable in the collection.",
        kind: "Method",
        signature: "Excel.PivotTableScopedCollection.getFirst => () => Excel.PivotTable",
        examples: [],
      },
      {
        name: "Excel.PivotTableScopedCollection.getFirstOrNullObject",
        description:
          "Gets the first PivotTable in the collection. The PivotTables in the collection are sorted top-to-bottom and left-to-right, such that the top-left table is the first PivotTable in the collection. If the collection is empty, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.PivotTableScopedCollection.getFirstOrNullObject => () => Excel.PivotTable",
        examples: [],
      },
      {
        name: "Excel.PivotTableScopedCollection.getItem",
        description: "Gets a PivotTable by name.",
        kind: "Method",
        signature: "Excel.PivotTableScopedCollection.getItem => (key: string) => Excel.PivotTable",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotTableStyle",
    apiList: [
      {
        name: "Excel.PivotTableStyle.name",
        description: "Specifies the name of the PivotTable style.",
        kind: "Property",
        signature: "Excel.PivotTableStyle.name: string",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyle.readOnly",
        description: "Specifies if this `PivotTableStyle` object is read-only.",
        kind: "Property",
        signature: "Excel.PivotTableStyle.readOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyle.delete",
        description: "Deletes the PivotTable style.",
        kind: "Method",
        signature: "Excel.PivotTableStyle.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyle.duplicate",
        description:
          "Creates a duplicate of this PivotTable style with copies of all the style elements.",
        kind: "Method",
        signature: "Excel.PivotTableStyle.duplicate => () => Excel.PivotTableStyle",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotTableStyleCollection",
    apiList: [
      {
        name: "Excel.PivotTableStyleCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.PivotTableStyleCollection.items: PivotTableStyle[]",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyleCollection.add",
        description: "Creates a blank `PivotTableStyle` with the specified name.",
        kind: "Method",
        signature:
          "Excel.PivotTableStyleCollection.add => (name: string, makeUniqueName?: boolean) => Excel.PivotTableStyle",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyleCollection.getCount",
        description: "Gets the number of PivotTable styles in the collection.",
        kind: "Method",
        signature:
          "Excel.PivotTableStyleCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyleCollection.getDefault",
        description: "Gets the default PivotTable style for the parent object's scope.",
        kind: "Method",
        signature: "Excel.PivotTableStyleCollection.getDefault => () => Excel.PivotTableStyle",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyleCollection.getItem",
        description: "Gets a `PivotTableStyle` by name.",
        kind: "Method",
        signature:
          "Excel.PivotTableStyleCollection.getItem => (name: string) => Excel.PivotTableStyle",
        examples: [],
      },
      {
        name: "Excel.PivotTableStyleCollection.setDefault",
        description: "Sets the default PivotTable style for use in the parent object's scope.",
        kind: "Method",
        signature:
          "Excel.PivotTableStyleCollection.setDefault => (newDefaultStyle: PivotTableStyle | string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PivotValueFilter",
    apiList: [
      {
        name: "Excel.PivotValueFilter.comparator",
        description:
          'The comparator is the static value to which other values are compared. The type of comparison is defined by the condition. For example, if comparator is "50" and condition is "greaterThan", all item values that are not greater than 50 will be removed by the filter.',
        kind: "Property",
        signature: "Excel.PivotValueFilter.comparator: number",
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.condition",
        description:
          "Specifies the condition for the filter, which defines the necessary filtering criteria.",
        kind: "Property",
        signature:
          'Excel.PivotValueFilter.condition: "Unknown" | "Equals" | "Between" | "GreaterThan" | "GreaterThanOrEqualTo" | "LessThan" | "LessThanOrEqualTo" | ValueFilterCondition | "TopN" | "BottomN"',
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.exclusive",
        description:
          "If `true`, filter *excludes* items that meet criteria. The default is `false` (filter to include items that meet criteria).",
        kind: "Property",
        signature: "Excel.PivotValueFilter.exclusive: boolean",
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.lowerBound",
        description: "The lower-bound of the range for the `between` filter condition.",
        kind: "Property",
        signature: "Excel.PivotValueFilter.lowerBound: number",
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.selectionType",
        description:
          "Specifies if the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.",
        kind: "Property",
        signature:
          'Excel.PivotValueFilter.selectionType: TopBottomSelectionType | "Items" | "Percent" | "Sum"',
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.threshold",
        description:
          'The "N" threshold number of items, percent, or sum to be filtered for a top/bottom filter condition.',
        kind: "Property",
        signature: "Excel.PivotValueFilter.threshold: number",
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.upperBound",
        description: "The upper-bound of the range for the `between` filter condition.",
        kind: "Property",
        signature: "Excel.PivotValueFilter.upperBound: number",
        examples: [],
      },
      {
        name: "Excel.PivotValueFilter.value",
        description: 'Name of the chosen "value" in the field by which to filter.',
        kind: "Property",
        signature: "Excel.PivotValueFilter.value: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PlaceholderErrorCellValue",
    apiList: [
      {
        name: "Excel.PlaceholderErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.PlaceholderErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.PlaceholderErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.PlaceholderErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.PlaceholderErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.PlaceholderErrorCellValue.errorType: ErrorCellValueType.placeholder | "Placeholder"',
        examples: [],
      },
      {
        name: "Excel.PlaceholderErrorCellValue.target",
        description:
          "`PlaceholderErrorCellValue` is used during processing, while data is downloaded. The `target` property represents the data that is downloading, the data for which the `PlaceholderErrorCellValue` object is a placeholder.",
        kind: "Property",
        signature:
          "Excel.PlaceholderErrorCellValue.target: LinkedEntityCellValue | WebImageCellValue",
        examples: [],
      },
      {
        name: "Excel.PlaceholderErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.PlaceholderErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.PresetCriteriaConditionalFormat",
    apiList: [
      {
        name: "Excel.PresetCriteriaConditionalFormat.format",
        description:
          "Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.",
        kind: "Property",
        signature: "Excel.PresetCriteriaConditionalFormat.format: Excel.ConditionalRangeFormat",
        examples: [
          'conditionalFormat.preset.format.font.color = "white";',
          'conditionalFormat.preset.format.font.color = "red";',
          'presetFormat.preset.format.font.color = "red";',
          "presetFormat.preset.format.font.bold = true;",
        ],
      },
      {
        name: "Excel.PresetCriteriaConditionalFormat.rule",
        description: "The rule of the conditional format.",
        kind: "Property",
        signature:
          "Excel.PresetCriteriaConditionalFormat.rule: Excel.ConditionalPresetCriteriaRule",
        examples: [
          "presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };",
          "conditionalFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage };",
        ],
      },
    ],
  },
  {
    objName: "Excel.Range",
    apiList: [
      {
        name: "Excel.Range.address",
        description:
          'Specifies the range reference in A1-style. Address value contains the sheet reference (e.g., "Sheet1!A1:B4").',
        kind: "Property",
        signature: "Excel.Range.address: string",
        examples: [
          'masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];',
          "`Copying the table headers spilled into ${spillRange.address}.`;",
          '`The address of the range B2:C5 is "${range.address}"`;',
          '`The address of the range "MyRange" is "${range.address}"`;',
          '`The address of the used range in the worksheet is "${range.address}"`;',
          '`The address of the entire worksheet range is "${range.address}"`;',
          '`The address of the selected range is "${selectedRange.address}"`;',
          "foundRange.address;",
          '"The active cell is " + activeCell.address;',
          '`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`;',
          "range.address;",
          'masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];',
          "cell.address;",
          "rangeEC.address;",
          "rangeER.address;",
          "tableDataRange.address;",
          "tableHeaderRange.address;",
          "activeTableRange.address;",
          "tableTotalsRange.address;",
          "dataBodyRange.address;",
          "headerRowRange.address;",
          "columnRange.address;",
          "totalRowRange.address;",
          "rowRange.address;",
          "selectedRange.address;",
          "usedRange.address;",
          '`The address of the frozen range (cells that are frozen in the top-and-left-most pane) is "${frozenRange.address}"`;',
        ],
      },
      {
        name: "Excel.Range.addressLocal",
        description:
          "Represents the range reference for the specified range in the language of the user.",
        kind: "Property",
        signature: "Excel.Range.addressLocal: string",
        examples: [],
      },
      {
        name: "Excel.Range.addressR1C1",
        description:
          'Specifies the range reference in R1C1-style. Address value contains the sheet reference (e.g., "Sheet1!R1C1:R4C2").',
        kind: "Property",
        signature: "Excel.Range.addressR1C1: string",
        examples: [],
      },
      {
        name: "Excel.Range.cellCount",
        description:
          "Specifies the number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647).",
        kind: "Property",
        signature: "Excel.Range.cellCount: number",
        examples: ["range.cellCount;"],
      },
      {
        name: "Excel.Range.columnCount",
        description: "Specifies the total number of columns in the range.",
        kind: "Property",
        signature: "Excel.Range.columnCount: number",
        examples: [
          "for (let j = 0; j < selectedRange.columnCount; j++) {\n      const cell = selectedRange.getCell(i, j);\n      cell.values = [[i * j]];\n\n      cell.untrack();\n    }",
          "const pasteToRange = activeWorksheet.getRangeByIndexes(\n    0,\n    usedRange.columnCount + 1,\n    expensesTableValues.length,\n    expensesTableValues[0].length\n  );",
        ],
      },
      {
        name: "Excel.Range.columnHidden",
        description:
          "Represents if all columns in the current range are hidden. Value is `true` when all columns in a range are hidden. Value is `false` when no columns in the range are hidden. Value is `null` when some columns in a range are hidden and other columns in the same range are not hidden.",
        kind: "Property",
        signature: "Excel.Range.columnHidden: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.columnIndex",
        description: "Specifies the column number of the first cell in the range. Zero-indexed.",
        kind: "Property",
        signature: "Excel.Range.columnIndex: number",
        examples: [],
      },
      {
        name: "Excel.Range.conditionalFormats",
        description: "The collection of `ConditionalFormats` that intersect the range.",
        kind: "Property",
        signature: "Excel.Range.conditionalFormats: Excel.ConditionalFormatCollection",
        examples: [
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);",
          'const conditionalFormat = range.conditionalFormats.getItemOrNullObject("0");',
          "const presetFormat = temperatureDataRange.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);",
          "const cellValueFormat = temperatureDataRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);",
          "range.conditionalFormats.clearAll();",
          "const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);",
          "const cfCount = range.conditionalFormats.getCount();",
          "const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);",
          "const conditionalFormat = activeTable.getDataBodyRange().conditionalFormats.add(Excel.ConditionalFormatType.custom);",
        ],
      },
      {
        name: "Excel.Range.dataValidation",
        description: "Returns a data validation object.",
        kind: "Property",
        signature: "Excel.Range.dataValidation: Excel.DataValidation",
        examples: [
          "commentsRange.dataValidation.clear();",
          "commentsRange.dataValidation.rule = redundantStringRule;",
          "rankingRange.dataValidation.clear();",
          "rankingRange.dataValidation.rule = greaterThanZeroRule;",
          "nameRange.dataValidation.clear();",
          "nameRange.dataValidation.rule = approvedListRule;",
        ],
      },
      {
        name: "Excel.Range.format",
        description:
          "Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.",
        kind: "Property",
        signature: "Excel.Range.format: Excel.RangeFormat",
        examples: [
          'headerRange.format.fill.color = "#4472C4";',
          'headerRange.format.font.color = "white";',
          "totalRange.format.font.bold = true;",
          'pinkColumnRange.format.fill.color = "pink";',
          "pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;",
          'range.format.fill.color = "#4472C4";',
          'range.format.font.color = "white";',
          "range.format.autofitColumns();",
          "activeWorksheet.getUsedRange().format.autofitColumns();",
          "activeWorksheet.getUsedRange().format.autofitRows();",
          'activeTable.getHeaderRowRange().format.fill.color = "#C70039";',
          'activeTable.getDataBodyRange().format.fill.color = "#DAF7A6";',
          'activeTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";',
          'activeTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";',
          'selectedRange.format.fill.color = "yellow";',
          "sumCell.format.autofitColumns();",
          "sheet.getUsedRange().format.autofitColumns();",
          "sheet.getUsedRange().format.autofitRows();",
          "cellRange.format.font.underline = Excel.RangeUnderlineStyle.none;",
          'cellRange.format.font.color = "#000000";',
          "activeTable.getRange().format.autofitColumns();",
          'chartTitle.format.horizontalAlignment = "Center";',
          "resultRange.format.autofitColumns();",
          "targetRange.format.autofitColumns();",
          'range.format.horizontalAlignment = "Right";',
          'range.format.borders.getItem("InsideHorizontal").style = "Continuous";',
          'range.format.borders.getItem("InsideVertical").style = "Continuous";',
          'range.format.borders.getItem("EdgeBottom").style = "Continuous";',
          'range.format.borders.getItem("EdgeLeft").style = "Continuous";',
          'range.format.borders.getItem("EdgeRight").style = "Continuous";',
          'range.format.borders.getItem("EdgeTop").style = "Continuous";',
          "const border = range.format.borders.getItem(Excel.BorderIndex.edgeTop);",
          "const border = range.format.borders.getItemAt(0);",
          "const rangeFill = range.format.fill;",
          "const rangeFont = range.format.font;",
          '[range.format.wrapText, range.format.fill.color, range.format.font.name].join("\\n");',
          "range.format.textOrientation = 90;",
          'range.format.verticalAlignment = "Justify";',
        ],
      },
      {
        name: "Excel.Range.formulas",
        description:
          "Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.",
        kind: "Property",
        signature: "Excel.Range.formulas: any[][]",
        examples: [
          "totalRange.formulas = totalFormulas;",
          'masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];',
          'targetCell.formulas = [["=A4:D4"]];',
          'range.formulas = [["=C3 * D3"]];',
          "range.formulas = data;",
          "JSON.stringify(range.formulas, null, 4);",
          'masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];',
          "range.formulas = formulas;",
        ],
      },
      {
        name: "Excel.Range.formulasLocal",
        description:
          'Represents the formula in A1-style notation, in the user\'s language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.',
        kind: "Property",
        signature: "Excel.Range.formulasLocal: any[][]",
        examples: [],
      },
      {
        name: "Excel.Range.formulasR1C1",
        description:
          "Represents the formula in R1C1-style notation. If a cell has no formula, its value is returned instead.",
        kind: "Property",
        signature: "Excel.Range.formulasR1C1: any[][]",
        examples: [],
      },
      {
        name: "Excel.Range.hasSpill",
        description:
          "Represents if all cells have a spill border. Returns `true` if all cells have a spill border, or `false` if all cells do not have a spill border. Returns `null` if there are cells both with and without spill borders within the range.",
        kind: "Property",
        signature: "Excel.Range.hasSpill: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.height",
        description:
          "Returns the distance in points, for 100% zoom, from the top edge of the range to the bottom edge of the range.",
        kind: "Property",
        signature: "Excel.Range.height: number",
        examples: [],
      },
      {
        name: "Excel.Range.hidden",
        description:
          "Represents if all cells in the current range are hidden. Value is `true` when all cells in a range are hidden. Value is `false` when no cells in the range are hidden. Value is `null` when some cells in a range are hidden and other cells in the same range are not hidden.",
        kind: "Property",
        signature: "Excel.Range.hidden: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.hyperlink",
        description: "Represents the hyperlink for the current range.",
        kind: "Property",
        signature: "Excel.Range.hyperlink: Excel.RangeHyperlink",
        examples: ["cellRange.hyperlink = hyperlink;"],
      },
      {
        name: "Excel.Range.isEntireColumn",
        description: "Represents if the current range is an entire column.",
        kind: "Property",
        signature: "Excel.Range.isEntireColumn: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.isEntireRow",
        description: "Represents if the current range is an entire row.",
        kind: "Property",
        signature: "Excel.Range.isEntireRow: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.left",
        description:
          "Returns the distance in points, for 100% zoom, from the left edge of the worksheet to the left edge of the range.",
        kind: "Property",
        signature: "Excel.Range.left: number",
        examples: [],
      },
      {
        name: "Excel.Range.numberFormat",
        description:
          "Represents Excel's number format code for the given range. For more information about Excel number formatting, see Number format codes.",
        kind: "Property",
        signature: "Excel.Range.numberFormat: any[][]",
        examples: [
          'totalRange.numberFormat = [["$0.00"]];',
          "range.numberFormat = formats;",
          "range.numberFormat = numberFormat;",
        ],
      },
      {
        name: "Excel.Range.numberFormatCategories",
        description: "Represents the category of number format of each cell.",
        kind: "Property",
        signature: "Excel.Range.numberFormatCategories: NumberFormatCategory[][]",
        examples: [],
      },
      {
        name: "Excel.Range.numberFormatLocal",
        description:
          "Represents Excel's number format code for the given range, based on the language settings of the user. Excel does not perform any language or format coercion when getting or setting the `numberFormatLocal` property. Any returned text uses the locally-formatted strings based on the language specified in the system settings.",
        kind: "Property",
        signature: "Excel.Range.numberFormatLocal: any[][]",
        examples: [],
      },
      {
        name: "Excel.Range.rowCount",
        description: "Returns the total number of rows in the range.",
        kind: "Property",
        signature: "Excel.Range.rowCount: number",
        examples: [
          "for (let i = 0; i < dataRange.rowCount; i++) {\n    const newSeries = bubbleChart.series.add(dataRange.values[i][0], i);\n    newSeries.setXAxisValues(dataRange.getCell(i, 1));\n    newSeries.setValues(dataRange.getCell(i, 2));\n    newSeries.setBubbleSizes(dataRange.getCell(i, 3));\n\n    newSeries.dataLabels.showSeriesName = true;\n    newSeries.dataLabels.showBubbleSize = true;\n    newSeries.dataLabels.showValue = false;\n  }",
          "for (let i = 0; i < selectedRange.rowCount; i++) {\n    for (let j = 0; j < selectedRange.columnCount; j++) {\n      const cell = selectedRange.getCell(i, j);\n      cell.values = [[i * j]];\n\n      cell.untrack();\n    }\n  }",
        ],
      },
      {
        name: "Excel.Range.rowHidden",
        description:
          "Represents if all rows in the current range are hidden. Value is `true` when all rows in a range are hidden. Value is `false` when no rows in the range are hidden. Value is `null` when some rows in a range are hidden and other rows in the same range are not hidden.",
        kind: "Property",
        signature: "Excel.Range.rowHidden: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.rowIndex",
        description: "Returns the row number of the first cell in the range. Zero-indexed.",
        kind: "Property",
        signature: "Excel.Range.rowIndex: number",
        examples: [],
      },
      {
        name: "Excel.Range.savedAsArray",
        description:
          "Represents if all the cells would be saved as an array formula. Returns `true` if all cells would be saved as an array formula, or `false` if all cells would not be saved as an array formula. Returns `null` if some cells would be saved as an array formula and some would not be.",
        kind: "Property",
        signature: "Excel.Range.savedAsArray: boolean",
        examples: [],
      },
      {
        name: "Excel.Range.sort",
        description: "Represents the range sort of the current range.",
        kind: "Property",
        signature: "Excel.Range.sort: Excel.RangeSort",
        examples: [
          "sortRange.sort.apply([\n    {\n      key: 3,\n      ascending: false,\n    },\n  ]);",
        ],
      },
      {
        name: "Excel.Range.style",
        description:
          "Represents the style of the current range. If the styles of the cells are inconsistent, `null` will be returned. For custom styles, the style name will be returned. For built-in styles, a string representing a value in the `BuiltInStyle` enum will be returned.",
        kind: "Property",
        signature: "Excel.Range.style: string",
        examples: [
          "range.style = Excel.BuiltInStyle.neutral;",
          'range.style = "Diagonal Orientation Style";',
        ],
      },
      {
        name: "Excel.Range.text",
        description:
          "Text values of the specified range. The text value will not depend on the cell width. The number sign (#) substitution that happens in the Excel UI will not affect the text value returned by the API.",
        kind: "Property",
        signature: "Excel.Range.text: string[][]",
        examples: ["JSON.stringify(range.text, null, 4);", "range.text;"],
      },
      {
        name: "Excel.Range.top",
        description:
          "Returns the distance in points, for 100% zoom, from the top edge of the worksheet to the top edge of the range.",
        kind: "Property",
        signature: "Excel.Range.top: number",
        examples: [],
      },
      {
        name: "Excel.Range.values",
        description:
          'Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string. If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.',
        kind: "Property",
        signature: "Excel.Range.values: any[][]",
        examples: [
          "headerRange.values = headers;",
          "dataRange.values = productData;",
          'activeWorksheet.getRange("F1").values = [["Moved Range"]];',
          "range.values = [[5]];",
          "range.values = data;",
          "JSON.stringify(range.values, null, 4);",
          'expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];',
          "let headerValues = headerRange.values;",
          "let bodyValues = bodyRange.values;",
          "let merchantColumnValues = columnRange.values;",
          'activeWorksheet.getRange("A11:A11").values = [["Results"]];',
          'activeWorksheet.getRange("A13:D13").values = headerValues;',
          'activeWorksheet.getRange("A14:D20").values = bodyValues;',
          'activeWorksheet.getRange("B23:B29").values = merchantColumnValues;',
          'activeWorksheet.getRange("A32:D32").values = secondRowValues;',
          "range.values = values;",
          '`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`;',
          'rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];',
          "rangeToSet.values = [[10, 20]];",
          '[rangeToGet.values, app.calculationMode, rangeToGet.values].join("\\n");',
          'table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];',
          'table.getDataBodyRange().getRow(1).values = [["D", 4]];',
          "const newSeries = bubbleChart.series.add(dataRange.values[i][0], i);",
          'expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];',
          'range.values = [[1], [20], [""], [5], ["test"]];',
          "const oldBigNumberString: string = bigNumberSource.values[0][0];",
          "resultRange.values = [[newBigNumberString]];",
          'activeWorksheet.getRange("F2").values = [["Copied Formula"]];',
          "let cellText = productsRange.values[i][0];",
          'activeWorksheet.getRange("F12").values = [["Moved Range:"]];',
          "cell.values = [[i * j]];",
          "const salesColumnValues = salesColumn.getDataBodyRange().values;",
          "const itemColumnValues = itemColumn.getDataBodyRange().values;",
          "salesColumn.getDataBodyRange().values = salesColumnValues;",
          "const yearColumnValues = yearColumn.getDataBodyRange().values;",
          "const voltageColumnValues = voltageColumn.getDataBodyRange().values;",
          "const reviewerColumnValues = reviewerColumn.getDataBodyRange().values;",
          "const bookColumnValues = bookColumn.getDataBodyRange().values;",
          "const authorColumnValues = authorColumn.getDataBodyRange().values;",
          "const ratingColumnValues = ratingColumn.getDataBodyRange().values;",
          "const expensesTableValues = activeTable.getRange().values;",
          "pasteToRange.values = expensesTableValues;",
          "newTable.getHeaderRowRange().values = activeTable.getHeaderRowRange().values;",
          "const tableDataBody = activeTable.getDataBodyRange().values;",
          "newTable.getHeaderRowRange().values = selectedRange.getRow(0).values;",
          "const tableDataBody = selectedRangeBody.values;",
          "const salesColumnValues = salesColumn.values;",
          "const ratingColumnValues = ratingColumn.values;",
        ],
      },
      {
        name: "Excel.Range.valueTypes",
        description: "Specifies the type of data in each cell.",
        kind: "Property",
        signature: "Excel.Range.valueTypes: RangeValueType[][]",
        examples: [],
      },
      {
        name: "Excel.Range.width",
        description:
          "Returns the distance in points, for 100% zoom, from the left edge of the range to the right edge of the range.",
        kind: "Property",
        signature: "Excel.Range.width: number",
        examples: [],
      },
      {
        name: "Excel.Range.worksheet",
        description: "The worksheet containing the current range.",
        kind: "Property",
        signature: "Excel.Range.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.Range.autoFill",
        description:
          "Fills a range from the current range to the destination range using the specified AutoFill logic. The destination range can be `null` or can extend the source range either horizontally or vertically. Discontiguous ranges are not supported. For more information, see Use AutoFill and Flash Fill.",
        kind: "Method",
        signature:
          "Excel.Range.autoFill(destinationRange?: string | Excel.Range, autoFillType?: Excel.AutoFillType): void",
        examples: [
          'sumCell.autoFill("K4:K7", Excel.AutoFillType.fillFormats);',
          'sumCell.autoFill("P4:P7", Excel.AutoFillType.fillCopy);',
        ],
      },
      {
        name: "Excel.Range.calculate",
        description: "Calculates a range of cells on a worksheet.",
        kind: "Method",
        signature: "Excel.Range.calculate => () => void",
        examples: [],
      },
      {
        name: "Excel.Range.clear",
        description: "Clear range values, format, fill, border, etc.",
        kind: "Method",
        signature: "Excel.Range.clear(applyTo?: Excel.ClearApplyTo): void",
        examples: ["range.clear();", "cellRange.clear(Excel.ClearApplyTo.hyperlinks);"],
      },
      {
        name: "Excel.Range.convertDataTypeToText",
        description: "Converts the range cells with data types into text.",
        kind: "Method",
        signature: "Excel.Range.convertDataTypeToText => () => void",
        examples: [],
      },
      {
        name: "Excel.Range.copyFrom",
        description:
          "Copies cell data or formatting from the source range or `RangeAreas` to the current range. The destination range can be a different size than the source range or `RangeAreas`. The destination is expanded automatically if it's smaller than the source. Note: Like the copy functionality in the Excel UI, if the destination range is an exact multiple greater than the source range in either rows or columns, then the source content is replicated multiple times. For example, a 2x2 range copy into a 2x6 range will result in 3 copies of the original 2x2 range.",
        kind: "Method",
        signature:
          "Excel.Range.copyFrom(sourceRange: string | Excel.Range | Excel.RangeAreas, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void",
        examples: [
          'activeWorksheet.getRange("G1").copyFrom("A1:E1");',
          'activeWorksheet.getRange("D1").copyFrom("A1:C1", Excel.RangeCopyType.all, true, false);',
          'activeWorksheet.getRange("D2").copyFrom("A2:C2", Excel.RangeCopyType.all, false, false);',
          'activeWorksheet.getRange("G2").copyFrom("A1:E1", Excel.RangeCopyType.formulas);',
        ],
      },
      {
        name: "Excel.Range.copyTo",
        description:
          "Copies cell data or formatting from the source range or `RangeAreas` to the current range. The destination range can be a different size than the source range or `RangeAreas`. The destination is expanded automatically if it's smaller than the source. Note: Like the copy functionality in the Excel UI, if the destination range is an exact multiple greater than the source range in either rows or columns, then the source content is replicated multiple times. For example, a 2x2 range copy into a 2x6 range will result in 3 copies of the original 2x2 range.",
        kind: "Method",
        signature:
          'Excel.Range.copyTo => { (sourceRange: string | RangeAreas | Range, copyType?: RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void; (sourceRange: string | RangeAreas | Range, copyType?: "All" | ... 3 more ... | "Link", skipBlanks?: boolean, transpose?: boolean): void; (sourceRange: Range | RangeAreas | string, copyType?: strin...',
        examples: [],
      },
      {
        name: "Excel.Range.delete",
        description: "Deletes the cells associated with the range.",
        kind: "Method",
        signature: "Excel.Range.delete(shift: Excel.DeleteShiftDirection): void",
        examples: ["range.delete(Excel.DeleteShiftDirection.up);", 'range.delete("Left");'],
      },
      {
        name: "Excel.Range.find",
        description:
          "Finds the given string based on the criteria specified. If the current range is larger than a single cell, then the search will be limited to that range, else the search will cover the entire sheet starting after that cell.",
        kind: "Method",
        signature: "Excel.Range.find(text: string, criteria: Excel.SearchCriteria) => Excel.Range",
        examples: [
          'let foundRange = activeTableRange.find("Food", {\n    completeMatch: true,\n    matchCase: false,\n    searchDirection: Excel.SearchDirection.forward,\n  });',
        ],
      },
      {
        name: "Excel.Range.findOrNullObject",
        description:
          "Finds the given string based on the criteria specified. If the current range is larger than a single cell, then the search will be limited to that range, else the search will cover the entire sheet starting after that cell. If there are no matches, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.Range.findOrNullObject => (text: string, criteria: Excel.SearchCriteria) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.flashFill",
        description:
          "Does a Flash Fill to the current range. Flash Fill automatically fills data when it senses a pattern, so the range must be a single column range and have data around it in order to find a pattern.",
        kind: "Method",
        signature: "Excel.Range.flashFill => () => void",
        examples: [],
      },
      {
        name: "Excel.Range.getAbsoluteResizedRange",
        description:
          "Gets a `Range` object with the same top-left cell as the current `Range` object, but with the specified numbers of rows and columns.",
        kind: "Method",
        signature:
          "Excel.Range.getAbsoluteResizedRange => (numRows: number, numColumns: number) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getBoundingRect",
        description:
          'Gets the smallest range object that encompasses the given ranges. For example, the `GetBoundingRect` of "B2:C5" and "D10:E15" is "B2:E15".',
        kind: "Method",
        signature: "Excel.Range.getBoundingRect(anotherRange: string | Excel.Range) => Excel.Range",
        examples: ['range = range.getBoundingRect("G4:H8");'],
      },
      {
        name: "Excel.Range.getCell",
        description:
          "Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.",
        kind: "Method",
        signature: "Excel.Range.getCell(row: number, column: number) => Excel.Range",
        examples: [
          "newSeries.setXAxisValues(dataRange.getCell(i, 1));",
          "newSeries.setValues(dataRange.getCell(i, 2));",
          "newSeries.setBubbleSizes(dataRange.getCell(i, 3));",
          "let cellRange = productsRange.getCell(i, 0);",
          "const cell = range.getCell(0, 0);",
          "const cell = selectedRange.getCell(i, j);",
        ],
      },
      {
        name: "Excel.Range.getCellProperties",
        description:
          "Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.",
        kind: "Method",
        signature:
          "Excel.Range.getCellProperties(cellPropertiesLoadOptions: Excel.CellPropertiesLoadOptions) => OfficeExtension.ClientResult<Excel.CellProperties[][]>",
        examples: [
          "const propertiesToGet = cell.getCellProperties({\n    address: true,\n    format: {\n      fill: {\n        color: true,\n      },\n      font: {\n        color: true,\n      },\n    },\n    style: true,\n  });",
        ],
      },
      {
        name: "Excel.Range.getColumn",
        description: "Gets a column contained in the range.",
        kind: "Method",
        signature: "Excel.Range.getColumn(column: number) => Excel.Range",
        examples: [
          "const range = activeWorksheet.getRange(rangeAddress).getColumn(1);",
          "const salesColumn = selectedRangeBody.getColumn(2);",
          "const ratingColumn = selectedRangeBody.getColumn(3);",
        ],
      },
      {
        name: "Excel.Range.getColumnProperties",
        description:
          "Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties. For properties that are not consistent across each cell within a given column, null will be returned.",
        kind: "Method",
        signature:
          "Excel.Range.getColumnProperties => (columnPropertiesLoadOptions: ColumnPropertiesLoadOptions) => OfficeExtension.ClientResult<ColumnProperties[]>",
        examples: [],
      },
      {
        name: "Excel.Range.getColumnsAfter",
        description: "Gets a certain number of columns to the right of the current `Range` object.",
        kind: "Method",
        signature: "Excel.Range.getColumnsAfter => (count?: number) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getColumnsBefore",
        description: "Gets a certain number of columns to the left of the current `Range` object.",
        kind: "Method",
        signature: "Excel.Range.getColumnsBefore => (count?: number) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getDataBodyRange",
        description: "Gets the range object associated with the data body of the rang.",
        kind: "Method",
        signature: "Excel.Range.getDataBodyRange() => Excel.Range",
        examples: ["const selectedRangeBody = selectedRange.getDataBodyRange();"],
      },
      {
        name: "Excel.Range.getDataClassificationIds",
        description:
          "Gets the data classification IDs for all PowerBI-based linked data types that are present in the range. 1st-party only.",
        kind: "Method",
        signature:
          "Excel.Range.getDataClassificationIds => () => OfficeExtension.ClientResult<string[][]>",
        examples: [],
      },
      {
        name: "Excel.Range.getDependents",
        description:
          "Returns a `WorkbookRangeAreas` object that represents the range containing all the dependents of a cell in the same worksheet or in multiple worksheets.",
        kind: "Method",
        signature: "Excel.Range.getDependents() => Excel.WorkbookRangeAreas",
        examples: [],
      },
      {
        name: "Excel.Range.getDirectDependents",
        description:
          "Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependent cells of a specified range in the same worksheet or across multiple worksheets.",
        kind: "Method",
        signature: "Excel.Range.getDirectDependents() => Excel.WorkbookRangeAreas",
        examples: [],
      },
      {
        name: "Excel.Range.getDirectPrecedents",
        description:
          "Returns a `WorkbookRangeAreas` object that represents the range containing all the direct precedent cells of a specified range in the same worksheet or across multiple worksheets.",
        kind: "Method",
        signature: "Excel.Range.getDirectPrecedents() => Excel.WorkbookRangeAreas",
        examples: [],
      },
      {
        name: "Excel.Range.getEntireColumn",
        description:
          'Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").',
        kind: "Method",
        signature: "Excel.Range.getEntireColumn() => Excel.Range",
        examples: ["const rangeEC = range.getEntireColumn();"],
      },
      {
        name: "Excel.Range.getEntireRow",
        description:
          'Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").',
        kind: "Method",
        signature: "Excel.Range.getEntireRow() => Excel.Range",
        examples: ["const rangeER = range.getEntireRow();"],
      },
      {
        name: "Excel.Range.getExtendedRange",
        description:
          "Returns a range object that includes the current range and up to the edge of the range, based on the provided direction. This matches the Ctrl+Shift+Arrow key behavior in the Excel on Windows UI.",
        kind: "Method",
        signature:
          "Excel.Range.getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: string | Excel.Range): Excel.Range",
        examples: [
          "let extendedRange = selectedRange.getExtendedRange(direction, activeCell);",
          "const extendedRange = selectedRange.getExtendedRange(direction, activeCell);",
        ],
      },
      {
        name: "Excel.Range.getImage",
        description:
          "Renders the range as a base64-encoded png image. *Important**: This API is currently unsupported in Excel for Mac. Visit OfficeDev/office-js Issue #235 for the current status.",
        kind: "Method",
        signature: "Excel.Range.getImage => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.Range.getIntersection",
        description:
          "Gets the range object that represents the rectangular intersection of the given ranges.",
        kind: "Method",
        signature: "Excel.Range.getIntersection(anotherRange: string | Excel.Range) => Excel.Range",
        examples: [
          'const range = activeWorksheet.getRange(rangeAddress).getIntersection("D4:G6");',
        ],
      },
      {
        name: "Excel.Range.getIntersectionOrNullObject",
        description:
          "Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.Range.getIntersectionOrNullObject => (anotherRange: Range | string) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getLastCell",
        description:
          'Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".',
        kind: "Method",
        signature: "Excel.Range.getLastCell() => Excel.Range",
        examples: ["const range = activeWorksheet.getRange(rangeAddress).getLastCell();"],
      },
      {
        name: "Excel.Range.getLastColumn",
        description:
          'Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".',
        kind: "Method",
        signature: "Excel.Range.getLastColumn() => Excel.Range",
        examples: ["const range = activeWorksheet.getRange(rangeAddress).getLastColumn();"],
      },
      {
        name: "Excel.Range.getLastRow",
        description:
          'Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".',
        kind: "Method",
        signature: "Excel.Range.getLastRow() => Excel.Range",
        examples: [
          "let grandTotalRange = range.getLastRow();",
          "const grandTotalRange = range.getLastRow();",
          "const range = activeWorksheet.getRange(rangeAddress).getLastRow();",
        ],
      },
      {
        name: "Excel.Range.getMergedAreas",
        description:
          "Returns a `RangeAreas` object that represents the merged areas in this range. Note that if the merged areas count in this range is more than 512, an `InvalidOperation` error will be thrown.",
        kind: "Method",
        signature: "Excel.Range.getMergedAreas => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.Range.getMergedAreasOrNullObject",
        description:
          "Returns a `RangeAreas` object that represents the merged areas in this range. Note that if the merged areas count in this range is more than 512, then this method will fail to return the result. If the `RangeAreas` object doesn't exist, then this function will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Range.getMergedAreasOrNullObject() => Excel.RangeAreas",
        examples: ["const mergedAreas = tableRange.getMergedAreasOrNullObject();"],
      },
      {
        name: "Excel.Range.getNumberFormatProperties",
        description:
          "Returns a collection of properties, each of which describe a characteristic of the selected number format.",
        kind: "Method",
        signature:
          "Excel.Range.getNumberFormatProperties => () => Excel.NumberFormatPropertyCollection",
        examples: [],
      },
      {
        name: "Excel.Range.getOffsetRange",
        description:
          "Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.",
        kind: "Method",
        signature:
          "Excel.Range.getOffsetRange(rowOffset: number, columnOffset: number) => Excel.Range",
        examples: [
          "const valueRange = dataRange.getOffsetRange(0, 1).getResizedRange(0, -1);",
          "const range = activeWorksheet.getRange(rangeAddress).getOffsetRange(-1, 4);",
        ],
      },
      {
        name: "Excel.Range.getPivotTables",
        description: "Gets a scoped collection of PivotTables that overlap with the range.",
        kind: "Method",
        signature:
          "Excel.Range.getPivotTables => (fullyContained?: boolean) => Excel.PivotTableScopedCollection",
        examples: [],
      },
      {
        name: "Excel.Range.getPrecedents",
        description:
          "Returns a `WorkbookRangeAreas` object that represents the range containing all the precedent cells of a specified range in the same worksheet or across multiple worksheets.",
        kind: "Method",
        signature: "Excel.Range.getPrecedents() => Excel.WorkbookRangeAreas",
        examples: [],
      },
      {
        name: "Excel.Range.getRangeEdge",
        description:
          "Returns a range object that is the edge cell of the data region that corresponds to the provided direction. This matches the Ctrl+Arrow key behavior in the Excel on Windows UI.",
        kind: "Method",
        signature:
          "Excel.Range.getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: string | Excel.Range): Excel.Range",
        examples: [
          "let rangeEdge = selectedRange.getRangeEdge(direction, activeCell);",
          "const rangeEdge = selectedRange.getRangeEdge(direction, activeCell);",
        ],
      },
      {
        name: "Excel.Range.getResizedRange",
        description:
          "Gets a `Range` object similar to the current `Range` object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.",
        kind: "Method",
        signature:
          "Excel.Range.getResizedRange(deltaRows: number, deltaColumns: number) => Excel.Range",
        examples: ["const valueRange = dataRange.getOffsetRange(0, 1).getResizedRange(0, -1);"],
      },
      {
        name: "Excel.Range.getRow",
        description: "Gets a row contained in the range.",
        kind: "Method",
        signature: "Excel.Range.getRow(row: number) => Excel.Range",
        examples: [
          'table.getDataBodyRange().getRow(1).values = [["D", 4]];',
          "const chartTitle = tableRange.getRow(0);",
          "const range = activeWorksheet.getRange(rangeAddress).getRow(1);",
          "newTable.getHeaderRowRange().values = selectedRange.getRow(0).values;",
        ],
      },
      {
        name: "Excel.Range.getRowProperties",
        description:
          "Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties. For properties that are not consistent across each cell within a given row, `null` will be returned.",
        kind: "Method",
        signature:
          "Excel.Range.getRowProperties => (rowPropertiesLoadOptions: RowPropertiesLoadOptions) => OfficeExtension.ClientResult<RowProperties[]>",
        examples: [],
      },
      {
        name: "Excel.Range.getRowsAbove",
        description: "Gets a certain number of rows above the current `Range` object.",
        kind: "Method",
        signature: "Excel.Range.getRowsAbove => (count?: number) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getRowsBelow",
        description: "Gets a certain number of rows below the current `Range` object.",
        kind: "Method",
        signature: "Excel.Range.getRowsBelow => (count?: number) => Excel.Range",
        examples: ['table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];'],
      },
      {
        name: "Excel.Range.getSpecialCells",
        description:
          "Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value. If no special cells are found, an `ItemNotFound` error will be thrown.",
        kind: "Method",
        signature:
          "Excel.Range.getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas",
        examples: [
          "let formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);",
          'const formulaRanges = usedRange.getSpecialCells("Constants", "LogicalText");',
          'const formulaRanges = usedRange.getSpecialCells("Formulas");',
        ],
      },
      {
        name: "Excel.Range.getSpecialCellsOrNullObject",
        description:
          "Gets the `RangeAreas` object, comprising one or more ranges, that represents all the cells that match the specified type and value. If no special cells are found, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          'Excel.Range.getSpecialCellsOrNullObject => { (cellType: SpecialCellType, cellValueType?: SpecialCellValueType): RangeAreas; (cellType: "Visible" | "Formulas" | "ConditionalFormats" | ... 4 more ... | "SameDataValidation", cellValueType?: "All" | ... 13 more ... | "Text"): RangeAreas; (cellType: string, cellValueType?: string): Excel.RangeAreas; }',
        examples: [],
      },
      {
        name: "Excel.Range.getSpillingToRange",
        description:
          "Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell.",
        kind: "Method",
        signature: "Excel.Range.getSpillingToRange() => Excel.Range",
        examples: [
          "let spillRange = targetCell.getSpillingToRange();",
          "const spillRange = targetCell.getSpillingToRange();",
        ],
      },
      {
        name: "Excel.Range.getSpillingToRangeOrNullObject",
        description:
          "Gets the range object containing the spill range when called on an anchor cell. If the range isn't an anchor cell or the spill range can't be found, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Range.getSpillingToRangeOrNullObject => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getSpillParent",
        description:
          "Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell.",
        kind: "Method",
        signature: "Excel.Range.getSpillParent => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getSpillParentOrNullObject",
        description:
          "Gets the range object containing the anchor cell for the cell getting spilled into. If it's not a spilled cell, or more than one cell is given, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Range.getSpillParentOrNullObject => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getSurroundingDataRegion",
        description:
          "Get the surrounding data region, as determined by Excel Ideas, as it relates to the current selection. The surrounding region is used by Excel Ideas to generate ideas that can be inserted into the workbook.",
        kind: "Method",
        signature: "Excel.Range.getSurroundingDataRegion => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getSurroundingRegion",
        description:
          "Returns a `Range` object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.",
        kind: "Method",
        signature: "Excel.Range.getSurroundingRegion => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getTables",
        description: "Gets a scoped collection of tables that overlap with the range.",
        kind: "Method",
        signature:
          "Excel.Range.getTables => (fullyContained?: boolean) => Excel.TableScopedCollection",
        examples: [],
      },
      {
        name: "Excel.Range.getUsedRange",
        description:
          "Returns the used range of the given range object. If there are no used cells within the range, this function will throw an `ItemNotFound` error.",
        kind: "Method",
        signature: "Excel.Range.getUsedRange => (valuesOnly?: boolean) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getUsedRangeOrNullObject",
        description:
          "Returns the used range of the given range object. If there are no used cells within the range, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Range.getUsedRangeOrNullObject => (valuesOnly?: boolean) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.getVisibleView",
        description: "Represents the visible rows of the current range.",
        kind: "Method",
        signature: "Excel.Range.getVisibleView() => Excel.RangeView",
        examples: ["let visibleRange = activeTable.getDataBodyRange().getVisibleView();"],
      },
      {
        name: "Excel.Range.group",
        description: "Groups columns and rows for an outline.",
        kind: "Method",
        signature: "Excel.Range.group(groupOption: Excel.GroupOption): void",
        examples: [
          'activeWorksheet.getRange("4:9").group(Excel.GroupOption.byRows);',
          'activeWorksheet.getRange("4:5").group(Excel.GroupOption.byRows);',
          'activeWorksheet.getRange("7:8").group(Excel.GroupOption.byRows);',
          'activeWorksheet.getRange("C:Q").group(Excel.GroupOption.byColumns);',
          'activeWorksheet.getRange("C:F").group(Excel.GroupOption.byColumns);',
          'activeWorksheet.getRange("H:K").group(Excel.GroupOption.byColumns);',
          'activeWorksheet.getRange("M:P").group(Excel.GroupOption.byColumns);',
        ],
      },
      {
        name: "Excel.Range.hideGroupDetails",
        description: "Hides the details of the row or column group.",
        kind: "Method",
        signature:
          'Excel.Range.hideGroupDetails => { (groupOption: GroupOption): void; (groupOption: "ByRows" | "ByColumns"): void; (groupOption: string): void; }',
        examples: [],
      },
      {
        name: "Excel.Range.insert",
        description:
          "Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new `Range` object at the now blank space.",
        kind: "Method",
        signature: "Excel.Range.insert(shift: Excel.InsertShiftDirection): Excel.Range",
        examples: ["range.insert(Excel.InsertShiftDirection.down);"],
      },
      {
        name: "Excel.Range.merge",
        description: "Merge the range cells into one region in the worksheet.",
        kind: "Method",
        signature: "Excel.Range.merge(across?: boolean) => void",
        examples: ["chartTitle.merge(true);", "range.merge(true);"],
      },
      {
        name: "Excel.Range.moveTo",
        description:
          "Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells. The destination range will be expanded automatically if it is smaller than the current range. Any cells in the destination range that are outside of the original range's area are not changed.",
        kind: "Method",
        signature: "Excel.Range.moveTo(destinationRange: string | Excel.Range) => void",
        examples: [
          'activeWorksheet.getRange("A1:E1").moveTo("G1");',
          'activeWorksheet.getRange("A1:E1").moveTo("G12");',
        ],
      },
      {
        name: "Excel.Range.removeDuplicates",
        description: "Removes duplicate values from the range specified by the columns.",
        kind: "Method",
        signature:
          "Excel.Range.removeDuplicates(columns: number[], includesHeader: boolean) => Excel.RemoveDuplicatesResult",
        examples: [
          "let deleteResult = range.removeDuplicates([0], true);",
          "const deleteResult = range.removeDuplicates([0], true);",
        ],
      },
      {
        name: "Excel.Range.replaceAll",
        description:
          "Finds and replaces the given string based on the criteria specified within the current range.",
        kind: "Method",
        signature:
          "Excel.Range.replaceAll => (text: string, replacement: string, criteria: Excel.ReplaceCriteria) => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.Range.select",
        description: "Selects the specified range in the Excel UI.",
        kind: "Method",
        signature: "Excel.Range.select() => void",
        examples: [
          "range.select();",
          "rangeEdge.select();",
          "extendedRange.select();",
          'activeWorksheet.getRange("B10:D14").select();',
        ],
      },
      {
        name: "Excel.Range.setCellProperties",
        description:
          "Updates the range based on a 2D array of cell properties, encapsulating things like font, fill, borders, and alignment.",
        kind: "Method",
        signature:
          "Excel.Range.setCellProperties(cellPropertiesData: Excel.SettableCellProperties[][]) => void",
        examples: [
          "range.setCellProperties([\n    [topHeaderProps, {}, {}, {}, {}],\n    [{}, {}, headerProps, headerProps, headerProps],\n    [{}, headerProps, nonApplicableProps, matchupScoreProps, matchupScoreProps],\n    [{}, headerProps, matchupScoreProps, nonApplicableProps, matchupScoreProps],\n    [{}, headerProps, matchupScoreProps, matchupScoreProps, nonApplicableProps],\n  ]);",
        ],
      },
      {
        name: "Excel.Range.setColumnProperties",
        description:
          "Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, and alignment.",
        kind: "Method",
        signature:
          "Excel.Range.setColumnProperties => (columnPropertiesData: SettableColumnProperties[]) => void",
        examples: [],
      },
      {
        name: "Excel.Range.setDirty",
        description: "Set a range to be recalculated when the next recalculation occurs.",
        kind: "Method",
        signature: "Excel.Range.setDirty => () => void",
        examples: [],
      },
      {
        name: "Excel.Range.setRowProperties",
        description:
          "Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, and alignment.",
        kind: "Method",
        signature:
          "Excel.Range.setRowProperties => (rowPropertiesData: SettableRowProperties[]) => void",
        examples: [],
      },
      {
        name: "Excel.Range.showCard",
        description: "Displays the card for an active cell if it has rich value content.",
        kind: "Method",
        signature: "Excel.Range.showCard => () => void",
        examples: [],
      },
      {
        name: "Excel.Range.showGroupDetails",
        description: "Shows the details of the row or column group.",
        kind: "Method",
        signature:
          'Excel.Range.showGroupDetails => { (groupOption: GroupOption): void; (groupOption: "ByRows" | "ByColumns"): void; (groupOption: string): void; }',
        examples: [],
      },
      {
        name: "Excel.Range.showTeachingCallout",
        description:
          "Shows a teaching callout next to the range. Title of the teaching callout.Body message of the teaching callout.",
        kind: "Method",
        signature: "Excel.Range.showTeachingCallout => (title: string, message: string) => void",
        examples: [],
      },
      {
        name: "Excel.Range.track",
        description:
          'Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created.',
        kind: "Method",
        signature: "Excel.Range.track => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Range.ungroup",
        description: "Ungroups columns and rows for an outline.",
        kind: "Method",
        signature:
          'Excel.Range.ungroup => { (groupOption: GroupOption): void; (groupOption: "ByRows" | "ByColumns"): void; (groupOption: string): void; }',
        examples: [],
      },
      {
        name: "Excel.Range.unmerge",
        description: "Unmerge the range cells into separate cells.",
        kind: "Method",
        signature: "Excel.Range.unmerge() => void",
        examples: ["range.unmerge();"],
      },
      {
        name: "Excel.Range.untrack",
        description:
          "Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.",
        kind: "Method",
        signature: "Excel.Range.untrack() => Excel.Range",
        examples: ["cell.untrack();"],
      },
    ],
  },
  {
    objName: "Excel.RangeAreas",
    apiList: [
      {
        name: "Excel.RangeAreas.address",
        description:
          'Returns the `RangeAreas` reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g., "Sheet1!A1:B4, Sheet1!D1:D4").',
        kind: "Property",
        signature: "Excel.RangeAreas.address: string",
        examples: [
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn, rangeAreas.address].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeAreas.addressLocal",
        description: "Returns the `RangeAreas` reference in the user locale.",
        kind: "Property",
        signature: "Excel.RangeAreas.addressLocal: string",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.addressR1C1",
        description:
          'Specifies the range reference in R1C1-style. Address value contains the sheet reference (e.g., "Sheet1!R1C1:R4C2"). Returns the `RangeAreas` reference in R1C1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g., "Sheet1!R1C1:R4C2, Sheet1!R1C4:R4C4").',
        kind: "Property",
        signature: "Excel.RangeAreas.addressR1C1: string",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.areaCount",
        description:
          "Returns the number of rectangular ranges that comprise this `RangeAreas` object.",
        kind: "Property",
        signature: "Excel.RangeAreas.areaCount: number",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.areas",
        description:
          "Returns a collection of rectangular ranges that comprise this `RangeAreas` object.",
        kind: "Property",
        signature: "Excel.RangeAreas.areas: Excel.RangeCollection",
        examples: ["const range = mergedAreas.areas.getItemAt(0);"],
      },
      {
        name: "Excel.RangeAreas.cellCount",
        description:
          "Returns the number of cells in the `RangeAreas` object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647).",
        kind: "Property",
        signature: "Excel.RangeAreas.cellCount: number",
        examples: [
          '[\n    `Address of the merged range: ${mergedAreas.address}`,\n    `Number of cells in the merged range: ${mergedAreas.cellCount}`,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeAreas.conditionalFormats",
        description:
          "Returns a collection of conditional formats that intersect with any cells in this `RangeAreas` object.",
        kind: "Property",
        signature: "Excel.RangeAreas.conditionalFormats: ConditionalFormatCollection",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.dataValidation",
        description: "Returns a data validation object for all ranges in the `RangeAreas`.",
        kind: "Property",
        signature: "Excel.RangeAreas.dataValidation: DataValidation",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.format",
        description:
          "Returns a `RangeFormat` object, encapsulating the font, fill, borders, alignment, and other properties for all ranges in the `RangeAreas` object.",
        kind: "Property",
        signature: "Excel.RangeAreas.format: Excel.RangeFormat",
        examples: [
          'rangeAreas.format.fill.color = "pink";',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn].join("\\n");',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn, rangeAreas.address].join("\\n");',
          'formulaRanges.format.fill.color = "pink";',
          'constantNumberRanges.format.fill.color = "pink";',
          'formulaLogicalNumberRanges.format.fill.color = "pink";',
          'foundRanges.format.fill.color = "green";',
          'formulaRanges.format.fill.color = "orange";',
          'formulaRanges.format.fill.color = "lightgreen";',
          'selectedRanges.format.fill.color = "lightblue";',
          'specifiedRanges.format.fill.color = "pink";',
        ],
      },
      {
        name: "Excel.RangeAreas.isEntireColumn",
        description:
          'Specifies if all the ranges on this `RangeAreas` object represent entire columns (e.g., "A:C, Q:Z").',
        kind: "Property",
        signature: "Excel.RangeAreas.isEntireColumn: boolean",
        examples: [
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn].join("\\n");',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn, rangeAreas.address].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeAreas.isEntireRow",
        description:
          'Specifies if all the ranges on this `RangeAreas` object represent entire rows (e.g., "1:3, 5:7").',
        kind: "Property",
        signature: "Excel.RangeAreas.isEntireRow: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.style",
        description:
          "Represents the style for all ranges in this `RangeAreas` object. If the styles of the cells are inconsistent, `null` will be returned. For custom styles, the style name will be returned. For built-in styles, a string representing a value in the `BuiltInStyle` enum will be returned.",
        kind: "Property",
        signature: "Excel.RangeAreas.style: string",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.worksheet",
        description: "Returns the worksheet for the current `RangeAreas`.",
        kind: "Property",
        signature: "Excel.RangeAreas.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.calculate",
        description: "Calculates all cells in the `RangeAreas`.",
        kind: "Method",
        signature: "Excel.RangeAreas.calculate => () => void",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.clear",
        description:
          "Clears values, format, fill, border, and other properties on each of the areas that comprise this `RangeAreas` object.",
        kind: "Method",
        signature:
          'Excel.RangeAreas.clear => { (applyTo?: ClearApplyTo): void; (applyTo?: "All" | "Formats" | "Contents" | "Hyperlinks" | "RemoveHyperlinks"): void; (applyTo?: string): void; }',
        examples: [],
      },
      {
        name: "Excel.RangeAreas.convertDataTypeToText",
        description: "Converts all cells in the `RangeAreas` with data types into text.",
        kind: "Method",
        signature: "Excel.RangeAreas.convertDataTypeToText => () => void",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.copyFrom",
        description:
          "Copies cell data or formatting from the source range or `RangeAreas` to the current `RangeAreas`. The destination `RangeAreas` can be a different size than the source range or `RangeAreas`. The destination will be expanded automatically if it is smaller than the source.",
        kind: "Method",
        signature:
          'Excel.RangeAreas.copyFrom => { (sourceRange: string | RangeAreas | Range, copyType?: RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void; (sourceRange: string | RangeAreas | Range, copyType?: "All" | ... 3 more ... | "Link", skipBlanks?: boolean, transpose?: boolean): void; (sourceRange: Range | RangeAreas | string, copyType?: strin...',
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getEntireColumn",
        description:
          'Returns a `RangeAreas` object that represents the entire columns of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11, H2", it returns a `RangeAreas` that represents columns "B:E, H:H").',
        kind: "Method",
        signature: "Excel.RangeAreas.getEntireColumn => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getEntireRow",
        description:
          'Returns a `RangeAreas` object that represents the entire rows of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11", it returns a `RangeAreas` that represents rows "4:11").',
        kind: "Method",
        signature: "Excel.RangeAreas.getEntireRow => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getIntersection",
        description:
          "Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`. If no intersection is found, an `ItemNotFound` error will be thrown.",
        kind: "Method",
        signature:
          "Excel.RangeAreas.getIntersection => (anotherRange: Range | RangeAreas | string) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getIntersectionOrNullObject",
        description:
          "Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`. If no intersection is found, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.RangeAreas.getIntersectionOrNullObject => (anotherRange: Range | RangeAreas | string) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getOffsetRangeAreas",
        description:
          "Returns a `RangeAreas` object that is shifted by the specific row and column offset. The dimension of the returned `RangeAreas` will match the original object. If the resulting `RangeAreas` is forced outside the bounds of the worksheet grid, an error will be thrown.",
        kind: "Method",
        signature:
          "Excel.RangeAreas.getOffsetRangeAreas => (rowOffset: number, columnOffset: number) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getSpecialCells",
        description:
          "Returns a `RangeAreas` object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.",
        kind: "Method",
        signature:
          'Excel.RangeAreas.getSpecialCells => { (cellType: SpecialCellType, cellValueType?: SpecialCellValueType): RangeAreas; (cellType: "Visible" | "Formulas" | "ConditionalFormats" | ... 4 more ... | "SameDataValidation", cellValueType?: "All" | ... 13 more ... | "Text"): RangeAreas; (cellType: string, cellValueType?: string): Excel.RangeAreas; }',
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getSpecialCellsOrNullObject",
        description:
          "Returns a `RangeAreas` object that represents all the cells that match the specified type and value. If no special cells are found that match the criteria, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          'Excel.RangeAreas.getSpecialCellsOrNullObject => { (cellType: SpecialCellType, cellValueType?: SpecialCellValueType): RangeAreas; (cellType: "Visible" | "Formulas" | "ConditionalFormats" | ... 4 more ... | "SameDataValidation", cellValueType?: "All" | ... 13 more ... | "Text"): RangeAreas; (cellType: string, cellValueType?: string): Excel.RangeAreas; }',
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getTables",
        description:
          "Returns a scoped collection of tables that overlap with any range in this `RangeAreas` object.",
        kind: "Method",
        signature:
          "Excel.RangeAreas.getTables => (fullyContained?: boolean) => Excel.TableScopedCollection",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getUsedRangeAreas",
        description:
          "Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object. If there are no used cells within the `RangeAreas`, the `ItemNotFound` error will be thrown.",
        kind: "Method",
        signature:
          "Excel.RangeAreas.getUsedRangeAreas => (valuesOnly?: boolean) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.getUsedRangeAreasOrNullObject",
        description:
          "Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object. If there are no used cells within the `RangeAreas`, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.RangeAreas.getUsedRangeAreasOrNullObject => (valuesOnly?: boolean) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.select",
        description: "Selects the specified range areas in the Excel UI.",
        kind: "Method",
        signature: "Excel.RangeAreas.select => () => void",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.setDirty",
        description: "Sets the `RangeAreas` to be recalculated when the next recalculation occurs.",
        kind: "Method",
        signature: "Excel.RangeAreas.setDirty => () => void",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.track",
        description:
          'Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created.',
        kind: "Method",
        signature: "Excel.RangeAreas.track => () => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.RangeAreas.untrack",
        description:
          "Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.",
        kind: "Method",
        signature: "Excel.RangeAreas.untrack => () => Excel.RangeAreas",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeAreasCollection",
    apiList: [
      {
        name: "Excel.RangeAreasCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.RangeAreasCollection.items: Excel.RangeAreas[]",
        examples: [],
      },
      {
        name: "Excel.RangeAreasCollection.getCount",
        description: "Gets the number of `RangeAreas` objects in this collection.",
        kind: "Method",
        signature:
          "Excel.RangeAreasCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.RangeAreasCollection.getItemAt",
        description: "Returns the `RangeAreas` object based on position in the collection.",
        kind: "Method",
        signature: "Excel.RangeAreasCollection.getItemAt => (index: number) => Excel.RangeAreas",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeBorder",
    apiList: [
      {
        name: "Excel.RangeBorder.color",
        description:
          'HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.RangeBorder.color: string",
        examples: [],
      },
      {
        name: "Excel.RangeBorder.sideIndex",
        description:
          "Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.",
        kind: "Property",
        signature:
          'Excel.RangeBorder.sideIndex: Excel.BorderIndex | "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal" | "DiagonalDown" | "DiagonalUp"',
        examples: ["border.sideIndex;"],
      },
      {
        name: "Excel.RangeBorder.style",
        description:
          "One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.RangeBorder.style: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot"',
        examples: [
          'range.format.borders.getItem("InsideHorizontal").style = "Continuous";',
          'range.format.borders.getItem("InsideVertical").style = "Continuous";',
          'range.format.borders.getItem("EdgeBottom").style = "Continuous";',
          'range.format.borders.getItem("EdgeLeft").style = "Continuous";',
          'range.format.borders.getItem("EdgeRight").style = "Continuous";',
          'range.format.borders.getItem("EdgeTop").style = "Continuous";',
          "border.style;",
        ],
      },
      {
        name: "Excel.RangeBorder.tintAndShade",
        description:
          "Specifies a double that lightens or darkens a color for the range border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color. A `null` value indicates that the border doesn't have a uniform `tintAndShade` setting.",
        kind: "Property",
        signature: "Excel.RangeBorder.tintAndShade: number",
        examples: [],
      },
      {
        name: "Excel.RangeBorder.weight",
        description:
          "Specifies the weight of the border around a range. See `Excel.BorderWeight` for details.",
        kind: "Property",
        signature:
          'Excel.RangeBorder.weight: BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeBorderCollection",
    apiList: [
      {
        name: "Excel.RangeBorderCollection.count",
        description: "Number of border objects in the collection.",
        kind: "Property",
        signature: "Excel.RangeBorderCollection.count: number",
        examples: [],
      },
      {
        name: "Excel.RangeBorderCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.RangeBorderCollection.items: RangeBorder[]",
        examples: [],
      },
      {
        name: "Excel.RangeBorderCollection.tintAndShade",
        description:
          "Specifies a double that lightens or darkens a color for range borders. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color. A `null` value indicates that the entire border collection doesn't have a uniform `tintAndShade` setting.",
        kind: "Property",
        signature: "Excel.RangeBorderCollection.tintAndShade: number",
        examples: [],
      },
      {
        name: "Excel.RangeBorderCollection.getItem",
        description: "Gets a border object using its name.",
        kind: "Method",
        signature:
          "Excel.RangeBorderCollection.getItem(index: Excel.BorderIndex): Excel.RangeBorder",
        examples: [
          'range.format.borders.getItem("InsideHorizontal").style = "Continuous";',
          'range.format.borders.getItem("InsideVertical").style = "Continuous";',
          'range.format.borders.getItem("EdgeBottom").style = "Continuous";',
          'range.format.borders.getItem("EdgeLeft").style = "Continuous";',
          'range.format.borders.getItem("EdgeRight").style = "Continuous";',
          'range.format.borders.getItem("EdgeTop").style = "Continuous";',
          "const border = range.format.borders.getItem(Excel.BorderIndex.edgeTop);",
        ],
      },
      {
        name: "Excel.RangeBorderCollection.getItemAt",
        description: "Gets a border object using its index.",
        kind: "Method",
        signature: "Excel.RangeBorderCollection.getItemAt(index: number) => Excel.RangeBorder",
        examples: ["const border = range.format.borders.getItemAt(0);"],
      },
    ],
  },
  {
    objName: "Excel.RangeCollection",
    apiList: [
      {
        name: "Excel.RangeCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.RangeCollection.items: Range[]",
        examples: [],
      },
      {
        name: "Excel.RangeCollection.getCount",
        description: "Returns the number of ranges in the `RangeCollection`.",
        kind: "Method",
        signature: "Excel.RangeCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.RangeCollection.getItemAt",
        description: "Returns the range object based on its position in the `RangeCollection`.",
        kind: "Method",
        signature: "Excel.RangeCollection.getItemAt(index: number) => Excel.Range",
        examples: ["const range = mergedAreas.areas.getItemAt(0);"],
      },
    ],
  },
  {
    objName: "Excel.RangeFill",
    apiList: [
      {
        name: "Excel.RangeFill.color",
        description:
          'HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")',
        kind: "Property",
        signature: "Excel.RangeFill.color: string",
        examples: [
          'headerRange.format.fill.color = "#4472C4";',
          'rangeAreas.format.fill.color = "pink";',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn].join("\\n");',
          'pinkColumnRange.format.fill.color = "pink";',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn, rangeAreas.address].join("\\n");',
          'range.format.fill.color = "#4472C4";',
          'formulaRanges.format.fill.color = "pink";',
          'constantNumberRanges.format.fill.color = "pink";',
          'formulaLogicalNumberRanges.format.fill.color = "pink";',
          'activeTable.getHeaderRowRange().format.fill.color = "#C70039";',
          'activeTable.getDataBodyRange().format.fill.color = "#DAF7A6";',
          'activeTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";',
          'activeTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";',
          'selectedRange.format.fill.color = "yellow";',
          'foundRanges.format.fill.color = "green";',
          'formulaRanges.format.fill.color = "orange";',
          'formulaRanges.format.fill.color = "lightgreen";',
          "rangeFill.color;",
          '[range.format.wrapText, range.format.fill.color, range.format.font.name].join("\\n");',
          'selectedRanges.format.fill.color = "lightblue";',
          'specifiedRanges.format.fill.color = "pink";',
        ],
      },
      {
        name: "Excel.RangeFill.pattern",
        description:
          "The pattern of a range. See `Excel.FillPattern` for details. LinearGradient and RectangularGradient are not supported. A `null` value indicates that the entire range doesn't have a uniform pattern setting.",
        kind: "Property",
        signature:
          'Excel.RangeFill.pattern: "None" | "Up" | "Down" | FillPattern | "Solid" | "Gray50" | "Gray75" | "Gray25" | "Horizontal" | "Vertical" | "Checker" | "SemiGray75" | "LightHorizontal" | "LightVertical" | ... 7 more ... | "RectangularGradient"',
        examples: [],
      },
      {
        name: "Excel.RangeFill.patternColor",
        description:
          'The HTML color code representing the color of the range pattern, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.RangeFill.patternColor: string",
        examples: [],
      },
      {
        name: "Excel.RangeFill.patternTintAndShade",
        description:
          "Specifies a double that lightens or darkens a pattern color for the range fill. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color. A `null` value indicates that the range doesn't have uniform `patternTintAndShade` settings.",
        kind: "Property",
        signature: "Excel.RangeFill.patternTintAndShade: number",
        examples: [],
      },
      {
        name: "Excel.RangeFill.tintAndShade",
        description:
          "Specifies a double that lightens or darkens a color for the range fill. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color. A `null` value indicates that the range doesn't have uniform `tintAndShade` settings.",
        kind: "Property",
        signature: "Excel.RangeFill.tintAndShade: number",
        examples: [],
      },
      {
        name: "Excel.RangeFill.clear",
        description: "Resets the range background.",
        kind: "Method",
        signature: "Excel.RangeFill.clear() => void",
        examples: ["rangeFill.clear();"],
      },
    ],
  },
  {
    objName: "Excel.RangeFont",
    apiList: [
      {
        name: "Excel.RangeFont.bold",
        description: "Represents the bold status of the font.",
        kind: "Property",
        signature: "Excel.RangeFont.bold: boolean",
        examples: ["totalRange.format.font.bold = true;"],
      },
      {
        name: "Excel.RangeFont.color",
        description:
          "HTML color code representation of the text color (e.g., #FF0000 represents Red).",
        kind: "Property",
        signature: "Excel.RangeFont.color: string",
        examples: [
          'headerRange.format.font.color = "white";',
          'range.format.font.color = "white";',
          'cellRange.format.font.color = "#000000";',
        ],
      },
      {
        name: "Excel.RangeFont.italic",
        description: "Specifies the italic status of the font.",
        kind: "Property",
        signature: "Excel.RangeFont.italic: boolean",
        examples: [
          '[\n    "Bold: " + style.font.bold,\n    "Font color: " + style.font.color,\n    "Italic: " + style.font.italic,\n    "Name: " + style.font.name,\n    "Size: " + style.font.size,\n    "Fill color: " + style.fill.color,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeFont.name",
        description:
          'Font name (e.g., "Calibri"). The name\'s length should not be greater than 31 characters.',
        kind: "Property",
        signature: "Excel.RangeFont.name: string",
        examples: [
          "rangeFont.name;",
          '[range.format.wrapText, range.format.fill.color, range.format.font.name].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeFont.size",
        description: "Font size.",
        kind: "Property",
        signature: "Excel.RangeFont.size: number",
        examples: [
          '[\n    "Bold: " + style.font.bold,\n    "Font color: " + style.font.color,\n    "Italic: " + style.font.italic,\n    "Name: " + style.font.name,\n    "Size: " + style.font.size,\n    "Fill color: " + style.fill.color,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeFont.strikethrough",
        description:
          "Specifies the strikethrough status of font. A `null` value indicates that the entire range doesn't have a uniform strikethrough setting.",
        kind: "Property",
        signature: "Excel.RangeFont.strikethrough: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFont.subscript",
        description:
          "Specifies the subscript status of font. Returns `true` if all the fonts of the range are subscript. Returns `false` if all the fonts of the range are superscript or normal (neither superscript, nor subscript). Returns `null` otherwise.",
        kind: "Property",
        signature: "Excel.RangeFont.subscript: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFont.superscript",
        description:
          "Specifies the superscript status of font. Returns `true` if all the fonts of the range are superscript. Returns `false` if all the fonts of the range are subscript or normal (neither superscript, nor subscript). Returns `null` otherwise.",
        kind: "Property",
        signature: "Excel.RangeFont.superscript: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFont.tintAndShade",
        description:
          "Specifies a double that lightens or darkens a color for the range font. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color. A `null` value indicates that the entire range doesn't have a uniform font `tintAndShade` setting.",
        kind: "Property",
        signature: "Excel.RangeFont.tintAndShade: number",
        examples: [],
      },
      {
        name: "Excel.RangeFont.underline",
        description:
          "Type of underline applied to the font. See `Excel.RangeUnderlineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.RangeFont.underline: Excel.RangeUnderlineStyle | "None" | "Single" | "Double" | "SingleAccountant" | "DoubleAccountant"',
        examples: ["cellRange.format.font.underline = Excel.RangeUnderlineStyle.none;"],
      },
    ],
  },
  {
    objName: "Excel.RangeFormat",
    apiList: [
      {
        name: "Excel.RangeFormat.autoIndent",
        description:
          "Specifies if text is automatically indented when text alignment is set to equal distribution.",
        kind: "Property",
        signature: "Excel.RangeFormat.autoIndent: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.borders",
        description: "Collection of border objects that apply to the overall range.",
        kind: "Property",
        signature: "Excel.RangeFormat.borders: Excel.RangeBorderCollection",
        examples: [
          'range.format.borders.getItem("InsideHorizontal").style = "Continuous";',
          'range.format.borders.getItem("InsideVertical").style = "Continuous";',
          'range.format.borders.getItem("EdgeBottom").style = "Continuous";',
          'range.format.borders.getItem("EdgeLeft").style = "Continuous";',
          'range.format.borders.getItem("EdgeRight").style = "Continuous";',
          'range.format.borders.getItem("EdgeTop").style = "Continuous";',
          "const border = range.format.borders.getItem(Excel.BorderIndex.edgeTop);",
          "const border = range.format.borders.getItemAt(0);",
        ],
      },
      {
        name: "Excel.RangeFormat.columnWidth",
        description:
          "Specifies the width of all colums within the range. If the column widths are not uniform, `null` will be returned.",
        kind: "Property",
        signature: "Excel.RangeFormat.columnWidth: number",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.fill",
        description: "Returns the fill object defined on the overall range.",
        kind: "Property",
        signature: "Excel.RangeFormat.fill: Excel.RangeFill",
        examples: [
          'headerRange.format.fill.color = "#4472C4";',
          'rangeAreas.format.fill.color = "pink";',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn].join("\\n");',
          'pinkColumnRange.format.fill.color = "pink";',
          '[rangeAreas.format.fill.color, rangeAreas.isEntireColumn, rangeAreas.address].join("\\n");',
          'range.format.fill.color = "#4472C4";',
          'formulaRanges.format.fill.color = "pink";',
          'constantNumberRanges.format.fill.color = "pink";',
          'formulaLogicalNumberRanges.format.fill.color = "pink";',
          'activeTable.getHeaderRowRange().format.fill.color = "#C70039";',
          'activeTable.getDataBodyRange().format.fill.color = "#DAF7A6";',
          'activeTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";',
          'activeTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";',
          'selectedRange.format.fill.color = "yellow";',
          'foundRanges.format.fill.color = "green";',
          'formulaRanges.format.fill.color = "orange";',
          'formulaRanges.format.fill.color = "lightgreen";',
          "const rangeFill = range.format.fill;",
          '[range.format.wrapText, range.format.fill.color, range.format.font.name].join("\\n");',
          'selectedRanges.format.fill.color = "lightblue";',
          'specifiedRanges.format.fill.color = "pink";',
        ],
      },
      {
        name: "Excel.RangeFormat.font",
        description: "Returns the font object defined on the overall range.",
        kind: "Property",
        signature: "Excel.RangeFormat.font: Excel.RangeFont",
        examples: [
          'headerRange.format.font.color = "white";',
          "totalRange.format.font.bold = true;",
          'range.format.font.color = "white";',
          "cellRange.format.font.underline = Excel.RangeUnderlineStyle.none;",
          'cellRange.format.font.color = "#000000";',
          "const rangeFont = range.format.font;",
          '[range.format.wrapText, range.format.fill.color, range.format.font.name].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeFormat.horizontalAlignment",
        description:
          "Represents the horizontal alignment for the specified object. See `Excel.HorizontalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.RangeFormat.horizontalAlignment: Excel.HorizontalAlignment | "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed"',
        examples: [
          "pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;",
          'chartTitle.format.horizontalAlignment = "Center";',
          'range.format.horizontalAlignment = "Right";',
        ],
      },
      {
        name: "Excel.RangeFormat.indentLevel",
        description: "An integer from 0 to 250 that indicates the indent level.",
        kind: "Property",
        signature: "Excel.RangeFormat.indentLevel: number",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.protection",
        description: "Returns the format protection object for a range.",
        kind: "Property",
        signature: "Excel.RangeFormat.protection: FormatProtection",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.readingOrder",
        description: "The reading order for the range.",
        kind: "Property",
        signature:
          'Excel.RangeFormat.readingOrder: ReadingOrder | "Context" | "LeftToRight" | "RightToLeft"',
        examples: [],
      },
      {
        name: "Excel.RangeFormat.rowHeight",
        description:
          "The height of all rows in the range. If the row heights are not uniform, `null` will be returned.",
        kind: "Property",
        signature: "Excel.RangeFormat.rowHeight: number",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.shrinkToFit",
        description:
          "Specifies if text automatically shrinks to fit in the available column width.",
        kind: "Property",
        signature: "Excel.RangeFormat.shrinkToFit: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.textOrientation",
        description:
          "The text orientation of all the cells within the range. The text orientation should be an integer either from -90 to 90, or 180 for vertically-oriented text. If the orientation within a range are not uniform, then `null` will be returned.",
        kind: "Property",
        signature: "Excel.RangeFormat.textOrientation: number",
        examples: ["range.format.textOrientation = 90;"],
      },
      {
        name: "Excel.RangeFormat.useStandardHeight",
        description:
          "Determines if the row height of the `Range` object equals the standard height of the sheet. Returns `true` if the row height of the `Range` object equals the standard height of the sheet. Returns `null` if the range contains more than one row and the rows aren't all the same height. Returns `false` otherwise. Note: This property is only intended to be set to `true`. Setting it to `false` has no effect.",
        kind: "Property",
        signature: "Excel.RangeFormat.useStandardHeight: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.useStandardWidth",
        description:
          "Specifies if the column width of the `Range` object equals the standard width of the sheet. Returns `true` if the column width of the `Range` object equals the standard width of the sheet. Returns `null` if the range contains more than one column and the columns aren't all the same height. Returns `false` otherwise. Note: This property is only intended to be set to `true`. Setting it to `false` has no effect.",
        kind: "Property",
        signature: "Excel.RangeFormat.useStandardWidth: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.verticalAlignment",
        description:
          "Represents the vertical alignment for the specified object. See `Excel.VerticalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.RangeFormat.verticalAlignment: Excel.VerticalAlignment | "Top" | "Center" | "Bottom" | "Justify" | "Distributed"',
        examples: ['range.format.verticalAlignment = "Justify";'],
      },
      {
        name: "Excel.RangeFormat.wrapText",
        description:
          "Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting",
        kind: "Property",
        signature: "Excel.RangeFormat.wrapText: boolean",
        examples: [
          '[range.format.wrapText, range.format.fill.color, range.format.font.name].join("\\n");',
        ],
      },
      {
        name: "Excel.RangeFormat.adjustIndent",
        description:
          "Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250 and is measured in characters.",
        kind: "Method",
        signature: "Excel.RangeFormat.adjustIndent => (amount: number) => void",
        examples: [],
      },
      {
        name: "Excel.RangeFormat.autofitColumns",
        description:
          "Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.",
        kind: "Method",
        signature: "Excel.RangeFormat.autofitColumns() => void",
        examples: [
          "range.format.autofitColumns();",
          "activeWorksheet.getUsedRange().format.autofitColumns();",
          "sumCell.format.autofitColumns();",
          "sheet.getUsedRange().format.autofitColumns();",
          "activeTable.getRange().format.autofitColumns();",
          "resultRange.format.autofitColumns();",
          "targetRange.format.autofitColumns();",
        ],
      },
      {
        name: "Excel.RangeFormat.autofitRows",
        description:
          "Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.",
        kind: "Method",
        signature: "Excel.RangeFormat.autofitRows() => void",
        examples: [
          "activeWorksheet.getUsedRange().format.autofitRows();",
          "sheet.getUsedRange().format.autofitRows();",
        ],
      },
    ],
  },
  {
    objName: "Excel.RangeHyperlink",
    apiList: [
      {
        name: "Excel.RangeHyperlink.address",
        description: "Represents the URL target for the hyperlink.",
        kind: "Property",
        signature: "Excel.RangeHyperlink.address: string",
        examples: [],
      },
      {
        name: "Excel.RangeHyperlink.documentReference",
        description: "Represents the document reference target for the hyperlink.",
        kind: "Property",
        signature: "Excel.RangeHyperlink.documentReference: string",
        examples: [],
      },
      {
        name: "Excel.RangeHyperlink.screenTip",
        description: "Represents the string displayed when hovering over the hyperlink.",
        kind: "Property",
        signature: "Excel.RangeHyperlink.screenTip: string",
        examples: [],
      },
      {
        name: "Excel.RangeHyperlink.textToDisplay",
        description:
          "Represents the string that is displayed in the top left most cell in the range.",
        kind: "Property",
        signature: "Excel.RangeHyperlink.textToDisplay: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeOptimization",
    apiList: [
      {
        name: "Excel.RangeOptimization.optimizationTypes",
        description: "The list of optimizations that can be applied to this range.",
        kind: "Property",
        signature: "Excel.RangeOptimization.optimizationTypes: RangeOptimizationType[]",
        examples: [],
      },
      {
        name: "Excel.RangeOptimization.range",
        description: "The address of a range that can be optimized.",
        kind: "Property",
        signature: "Excel.RangeOptimization.range: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeOptimizationCollection",
    apiList: [
      {
        name: "Excel.RangeOptimizationCollection.allocatedCells",
        description:
          "The number of cells that are allocated in the worksheet associated with this collection.",
        kind: "Property",
        signature: "Excel.RangeOptimizationCollection.allocatedCells: number",
        examples: [],
      },
      {
        name: "Excel.RangeOptimizationCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.RangeOptimizationCollection.items: RangeOptimization[]",
        examples: [],
      },
      {
        name: "Excel.RangeOptimizationCollection.optimizableCells",
        description: "The number of cells in the collection that can be optimized.",
        kind: "Property",
        signature: "Excel.RangeOptimizationCollection.optimizableCells: number",
        examples: [],
      },
      {
        name: "Excel.RangeOptimizationCollection.getCount",
        description: "Returns the number of ranges in the collection.",
        kind: "Method",
        signature:
          "Excel.RangeOptimizationCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.RangeOptimizationCollection.getItemAt",
        description: "Returns a range optimization by its index in the collection.",
        kind: "Method",
        signature:
          "Excel.RangeOptimizationCollection.getItemAt => (index: number) => Excel.RangeOptimization",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeReference",
    apiList: [
      {
        name: "Excel.RangeReference.address",
        description: 'The address of the range, for example "SheetName!A1:B5".',
        kind: "Property",
        signature: "Excel.RangeReference.address: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeSort",
    apiList: [
      {
        name: "Excel.RangeSort.apply",
        description: "Perform a sort operation.",
        kind: "Method",
        signature:
          "Excel.RangeSort.apply(fields: Excel.SortField[], matchCase?: boolean, hasHeaders?: boolean, orientation?: Excel.SortOrientation, method?: Excel.SortMethod): void",
        examples: [
          "sortRange.sort.apply([\n    {\n      key: 3,\n      ascending: false,\n    },\n  ]);",
        ],
      },
    ],
  },
  {
    objName: "Excel.RangeValuesPreview",
    apiList: [
      {
        name: "Excel.RangeValuesPreview.dismiss",
        description: "Dismisses the preview.",
        kind: "Method",
        signature: "Excel.RangeValuesPreview.dismiss => () => void",
        examples: [],
      },
      {
        name: "Excel.RangeValuesPreview.registerEventDismissed",
        description: "Register Event dismissed",
        kind: "Method",
        signature: "Excel.RangeValuesPreview.registerEventDismissed => () => void",
        examples: [],
      },
      {
        name: "Excel.RangeValuesPreview.show",
        description:
          "Shows the preview of values in the range. The range dimensions are defined by the anchor cell and dimensions of the values array.",
        kind: "Method",
        signature:
          "Excel.RangeValuesPreview.show => (anchorCellAddress: string, values: string[][], options?: Excel.RangeValuesPreviewOptions) => void",
        examples: [],
      },
      {
        name: "Excel.RangeValuesPreview.unregisterEventDismissed",
        description: "Register Event dismissed",
        kind: "Method",
        signature: "Excel.RangeValuesPreview.unregisterEventDismissed => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeValuesPreviewOptions",
    apiList: [
      {
        name: "Excel.RangeValuesPreviewOptions.autoexpandTable",
        description:
          "Determines whether the range values preview autoexpands an adjacent table, if any.",
        kind: "Property",
        signature: "Excel.RangeValuesPreviewOptions.autoexpandTable: boolean",
        examples: [],
      },
      {
        name: "Excel.RangeValuesPreviewOptions.autofitColumns",
        description: "Determines whether the range values preview autofits columns.",
        kind: "Property",
        signature: "Excel.RangeValuesPreviewOptions.autofitColumns: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeView",
    apiList: [
      {
        name: "Excel.RangeView.cellAddresses",
        description: "Represents the cell addresses of the `RangeView`.",
        kind: "Property",
        signature: "Excel.RangeView.cellAddresses: any[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.columnCount",
        description: "The number of visible columns.",
        kind: "Property",
        signature: "Excel.RangeView.columnCount: number",
        examples: [],
      },
      {
        name: "Excel.RangeView.formulas",
        description:
          "Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.",
        kind: "Property",
        signature: "Excel.RangeView.formulas: any[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.formulasLocal",
        description:
          'Represents the formula in A1-style notation, in the user\'s language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.',
        kind: "Property",
        signature: "Excel.RangeView.formulasLocal: any[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.formulasR1C1",
        description:
          "Represents the formula in R1C1-style notation. If a cell has no formula, its value is returned instead.",
        kind: "Property",
        signature: "Excel.RangeView.formulasR1C1: any[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.index",
        description: "Returns a value that represents the index of the `RangeView`.",
        kind: "Property",
        signature: "Excel.RangeView.index: number",
        examples: [],
      },
      {
        name: "Excel.RangeView.numberFormat",
        description: "Represents Excel's number format code for the given cell.",
        kind: "Property",
        signature: "Excel.RangeView.numberFormat: any[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.rowCount",
        description: "The number of visible rows.",
        kind: "Property",
        signature: "Excel.RangeView.rowCount: number",
        examples: [],
      },
      {
        name: "Excel.RangeView.rows",
        description: "Represents a collection of range views associated with the range.",
        kind: "Property",
        signature: "Excel.RangeView.rows: RangeViewCollection",
        examples: [],
      },
      {
        name: "Excel.RangeView.text",
        description:
          "Text values of the specified range. The text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API.",
        kind: "Property",
        signature: "Excel.RangeView.text: string[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.values",
        description:
          "Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.",
        kind: "Property",
        signature: "Excel.RangeView.values: any[][]",
        examples: ["visibleRange.values;"],
      },
      {
        name: "Excel.RangeView.valueTypes",
        description: "Represents the type of data of each cell.",
        kind: "Property",
        signature: "Excel.RangeView.valueTypes: RangeValueType[][]",
        examples: [],
      },
      {
        name: "Excel.RangeView.getRange",
        description: "Gets the parent range associated with the current `RangeView`.",
        kind: "Method",
        signature: "Excel.RangeView.getRange => () => Excel.Range",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RangeViewCollection",
    apiList: [
      {
        name: "Excel.RangeViewCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.RangeViewCollection.items: RangeView[]",
        examples: [],
      },
      {
        name: "Excel.RangeViewCollection.getCount",
        description: "Gets the number of `RangeView` objects in the collection.",
        kind: "Method",
        signature:
          "Excel.RangeViewCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.RangeViewCollection.getItemAt",
        description: "Gets a `RangeView` row via its index. Zero-indexed.",
        kind: "Method",
        signature: "Excel.RangeViewCollection.getItemAt => (index: number) => Excel.RangeView",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ReferenceCellValue",
    apiList: [
      {
        name: "Excel.ReferenceCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature:
          'Excel.ReferenceCellValue.basicType: RangeValueType | "Error" | "Boolean" | "Double" | "Empty" | "String"',
        examples: [],
      },
      {
        name: "Excel.ReferenceCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: "Excel.ReferenceCellValue.basicValue: string | number | boolean",
        examples: [],
      },
      {
        name: "Excel.ReferenceCellValue.reference",
        description:
          "Represents the index into the `referencedValues` properties of cell values such as `EntityCellValue` and `ArrayCellValue`.",
        kind: "Property",
        signature: "Excel.ReferenceCellValue.reference: number",
        examples: [],
      },
      {
        name: "Excel.ReferenceCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.ReferenceCellValue.type: CellValueType.reference | "Reference"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RefErrorCellValue",
    apiList: [
      {
        name: "Excel.RefErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.RefErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.RefErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.RefErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.RefErrorCellValue.errorSubType",
        description: "Represents the type of `RefErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.RefErrorCellValue.errorSubType: "Unknown" | RefErrorCellValueSubType | "ExternalLinksStructuredRef" | "ExternalLinksCalculatedRef"',
        examples: [],
      },
      {
        name: "Excel.RefErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.RefErrorCellValue.errorType: ErrorCellValueType.ref | "Ref"',
        examples: [],
      },
      {
        name: "Excel.RefErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.RefErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RemoveDuplicatesResult",
    apiList: [
      {
        name: "Excel.RemoveDuplicatesResult.removed",
        description: "Number of duplicated rows removed by the operation.",
        kind: "Property",
        signature: "Excel.RemoveDuplicatesResult.removed: number",
        examples: [
          '[\n    deleteResult.removed + " entries with duplicate names removed.",\n    deleteResult.uniqueRemaining + " entries with unique names remain in the range.",\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.RemoveDuplicatesResult.uniqueRemaining",
        description: "Number of remaining unique rows present in the resulting range.",
        kind: "Property",
        signature: "Excel.RemoveDuplicatesResult.uniqueRemaining: number",
        examples: [
          '[\n    deleteResult.removed + " entries with duplicate names removed.",\n    deleteResult.uniqueRemaining + " entries with unique names remain in the range.",\n  ].join("\\n");',
        ],
      },
    ],
  },
  {
    objName: "Excel.ReplaceCriteria",
    apiList: [
      {
        name: "Excel.ReplaceCriteria.completeMatch",
        description:
          "Specifies if the match needs to be complete or partial. A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`). Default is `false` (partial).",
        kind: "Property",
        signature: "Excel.ReplaceCriteria.completeMatch: boolean",
        examples: [],
      },
      {
        name: "Excel.ReplaceCriteria.matchCase",
        description:
          "Specifies if the match is case-sensitive. Default is `false` (case-insensitive).",
        kind: "Property",
        signature: "Excel.ReplaceCriteria.matchCase: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RowColumnPivotHierarchy",
    apiList: [
      {
        name: "Excel.RowColumnPivotHierarchy.fields",
        description: "Returns the PivotFields associated with the RowColumnPivotHierarchy.",
        kind: "Property",
        signature: "Excel.RowColumnPivotHierarchy.fields: Excel.PivotFieldCollection",
        examples: [
          'let filterField = dateHierarchy.fields.getItem("Date Updated");',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'const filterField = dateHierarchy.fields.getItem("Date Updated");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
      {
        name: "Excel.RowColumnPivotHierarchy.id",
        description: "ID of the RowColumnPivotHierarchy.",
        kind: "Property",
        signature: "Excel.RowColumnPivotHierarchy.id: string",
        examples: [],
      },
      {
        name: "Excel.RowColumnPivotHierarchy.name",
        description: "Name of the RowColumnPivotHierarchy.",
        kind: "Property",
        signature: "Excel.RowColumnPivotHierarchy.name: string",
        examples: [],
      },
      {
        name: "Excel.RowColumnPivotHierarchy.position",
        description: "Position of the RowColumnPivotHierarchy.",
        kind: "Property",
        signature: "Excel.RowColumnPivotHierarchy.position: number",
        examples: [],
      },
      {
        name: "Excel.RowColumnPivotHierarchy.setToDefault",
        description: "Reset the RowColumnPivotHierarchy back to its default values.",
        kind: "Method",
        signature: "Excel.RowColumnPivotHierarchy.setToDefault => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.RowColumnPivotHierarchyCollection",
    apiList: [
      {
        name: "Excel.RowColumnPivotHierarchyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.RowColumnPivotHierarchyCollection.items: RowColumnPivotHierarchy[]",
        examples: [],
      },
      {
        name: "Excel.RowColumnPivotHierarchyCollection.add",
        description:
          "Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column, or filter axis, it will be removed from that location.",
        kind: "Method",
        signature:
          "Excel.RowColumnPivotHierarchyCollection.add(pivotHierarchy: Excel.PivotHierarchy) => Excel.RowColumnPivotHierarchy",
        examples: [
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));',
          'pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));',
          'dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));',
        ],
      },
      {
        name: "Excel.RowColumnPivotHierarchyCollection.getCount",
        description: "Gets the number of pivot hierarchies in the collection.",
        kind: "Method",
        signature:
          "Excel.RowColumnPivotHierarchyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.RowColumnPivotHierarchyCollection.getItem",
        description: "Gets a RowColumnPivotHierarchy by its name or ID.",
        kind: "Method",
        signature:
          "Excel.RowColumnPivotHierarchyCollection.getItem(name: string) => Excel.RowColumnPivotHierarchy",
        examples: [
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
      {
        name: "Excel.RowColumnPivotHierarchyCollection.remove",
        description: "Removes the PivotHierarchy from the current axis.",
        kind: "Method",
        signature:
          "Excel.RowColumnPivotHierarchyCollection.remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy) => void",
        examples: ["pivotTable.columnHierarchies.remove(column);"],
      },
    ],
  },
  {
    objName: "Excel.RowPropertiesLoadOptions",
    apiList: [
      {
        name: "Excel.RowPropertiesLoadOptions.address",
        description: "Specifies whether to load on the `address` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.address: boolean",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.addressLocal",
        description:
          "Specifies whether to load on the `addressLocal` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.addressLocal: boolean",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.format",
        description: "Specifies whether to load on the `format` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature:
          "Excel.RowPropertiesLoadOptions.format: CellPropertiesFormatLoadOptions & { rowHeight?: boolean; }",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.hidden",
        description: "Specifies whether to load on the `hidden` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.hidden: boolean",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.hyperlink",
        description:
          "Specifies whether to load on the `hyperlink` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.hyperlink: boolean",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.rowHidden",
        description:
          "Specifies whether to load on the `rowHidden` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.rowHidden: boolean",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.rowIndex",
        description:
          "Specifies whether to load on the `rowIndex` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.rowIndex: boolean",
        examples: [],
      },
      {
        name: "Excel.RowPropertiesLoadOptions.style",
        description: "Specifies whether to load on the `style` property. [Api set: ExcelApi 1.9]",
        kind: "Property",
        signature: "Excel.RowPropertiesLoadOptions.style: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SearchCriteria",
    apiList: [
      {
        name: "Excel.SearchCriteria.completeMatch",
        description:
          "Specifies if the match needs to be complete or partial. A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`). Default is `false` (partial).",
        kind: "Property",
        signature: "Excel.SearchCriteria.completeMatch: boolean",
        examples: [],
      },
      {
        name: "Excel.SearchCriteria.matchCase",
        description:
          "Specifies if the match is case-sensitive. Default is `false` (case-insensitive).",
        kind: "Property",
        signature: "Excel.SearchCriteria.matchCase: boolean",
        examples: [],
      },
      {
        name: "Excel.SearchCriteria.searchDirection",
        description:
          "Specifies the search direction. Default is forward. See `Excel.SearchDirection`.",
        kind: "Property",
        signature:
          'Excel.SearchCriteria.searchDirection: SearchDirection | "Forward" | "Backwards"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Setting",
    apiList: [
      {
        name: "Excel.Setting.key",
        description: "The key that represents the ID of the setting.",
        kind: "Property",
        signature: "Excel.Setting.key: string",
        examples: [],
      },
      {
        name: "Excel.Setting.value",
        description: "Represents the value stored for this setting.",
        kind: "Property",
        signature: "Excel.Setting.value: any",
        examples: ['"Workbook needs review : " + needsReview.value;'],
      },
      {
        name: "Excel.Setting.delete",
        description: "Deletes the setting.",
        kind: "Method",
        signature: "Excel.Setting.delete => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SettingCollection",
    apiList: [
      {
        name: "Excel.SettingCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.SettingCollection.items: Setting[]",
        examples: [],
      },
      {
        name: "Excel.SettingCollection.add",
        description: "Sets or adds the specified setting to the workbook.",
        kind: "Method",
        signature: "Excel.SettingCollection.add(key: string, value: any) => Excel.Setting",
        examples: ['settings.add("NeedsReview", true);'],
      },
      {
        name: "Excel.SettingCollection.getCount",
        description: "Gets the number of settings in the collection.",
        kind: "Method",
        signature: "Excel.SettingCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.SettingCollection.getItem",
        description: "Gets a setting entry via the key.",
        kind: "Method",
        signature: "Excel.SettingCollection.getItem(key: string) => Excel.Setting",
        examples: ['let needsReview = settings.getItem("NeedsReview");'],
      },
    ],
  },
  {
    objName: "Excel.Shape",
    apiList: [
      {
        name: "Excel.Shape.altTextDescription",
        description: "Specifies the alternative description text for a `Shape` object.",
        kind: "Property",
        signature: "Excel.Shape.altTextDescription: string",
        examples: [],
      },
      {
        name: "Excel.Shape.altTextTitle",
        description: "Specifies the alternative title text for a `Shape` object.",
        kind: "Property",
        signature: "Excel.Shape.altTextTitle: string",
        examples: [],
      },
      {
        name: "Excel.Shape.connectionSiteCount",
        description: "Returns the number of connection sites on this shape.",
        kind: "Property",
        signature: "Excel.Shape.connectionSiteCount: number",
        examples: [],
      },
      {
        name: "Excel.Shape.displayName",
        description:
          "Gets the display name of the shape. A newly created shape has a generated name that is localized and may not match its `name`. In this scenario, you can use this API to get the name that is displayed in the UI.",
        kind: "Property",
        signature: "Excel.Shape.displayName: string",
        examples: [],
      },
      {
        name: "Excel.Shape.fill",
        description: "Returns the fill formatting of this shape.",
        kind: "Property",
        signature: "Excel.Shape.fill: Excel.ShapeFill",
        examples: ['shape.fill.foregroundColor = "yellow";', "shape.fill.clear();"],
      },
      {
        name: "Excel.Shape.geometricShape",
        description:
          'Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".',
        kind: "Property",
        signature: "Excel.Shape.geometricShape: GeometricShape",
        examples: [],
      },
      {
        name: "Excel.Shape.geometricShapeType",
        description:
          'Specifies the geometric shape type of this geometric shape. See `Excel.GeometricShapeType` for details. Returns `null` if the shape type is not "GeometricShape".',
        kind: "Property",
        signature:
          'Excel.Shape.geometricShapeType: "Cube" | "Pie" | "Funnel" | "Diamond" | "Triangle" | "Plus" | "Corner" | "Donut" | GeometricShapeType | "LineInverse" | "RightTriangle" | "Rectangle" | "Parallelogram" | ... 164 more ... | "ChartPlus"',
        examples: [],
      },
      {
        name: "Excel.Shape.group",
        description:
          'Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".',
        kind: "Property",
        signature: "Excel.Shape.group: Excel.ShapeGroup",
        examples: ['const shapeGroup = activeWorksheet.shapes.getItem("Group").group;'],
      },
      {
        name: "Excel.Shape.height",
        description:
          "Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value or zero as an input.",
        kind: "Property",
        signature: "Excel.Shape.height: number",
        examples: [
          "shape.height = 175;",
          "shape.height = 100;",
          "shape.height = 150;",
          "textbox.height = 20;",
        ],
      },
      {
        name: "Excel.Shape.id",
        description: "Specifies the shape identifier.",
        kind: "Property",
        signature: "Excel.Shape.id: string",
        examples: [],
      },
      {
        name: "Excel.Shape.image",
        description:
          'Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".',
        kind: "Property",
        signature: "Excel.Shape.image: Excel.Image",
        examples: ['const image = activeWorksheet.shapes.getItem("Image").image;'],
      },
      {
        name: "Excel.Shape.left",
        description:
          "The distance, in points, from the left side of the shape to the left side of the worksheet. Throws an `InvalidArgument` exception when set with a negative value as an input.",
        kind: "Property",
        signature: "Excel.Shape.left: number",
        examples: [
          "shape.left = 5;",
          "shape.left = 300;",
          "shape.left = 100;",
          "textbox.left = 100;",
        ],
      },
      {
        name: "Excel.Shape.level",
        description:
          "Specifies the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.",
        kind: "Property",
        signature: "Excel.Shape.level: number",
        examples: [],
      },
      {
        name: "Excel.Shape.line",
        description:
          'Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".',
        kind: "Property",
        signature: "Excel.Shape.line: Excel.Line",
        examples: ['const line = shapes.getItem("StraightLine").line;'],
      },
      {
        name: "Excel.Shape.lineFormat",
        description: "Returns the line formatting of this shape.",
        kind: "Property",
        signature: "Excel.Shape.lineFormat: ShapeLineFormat",
        examples: [],
      },
      {
        name: "Excel.Shape.lockAspectRatio",
        description: "Specifies if the aspect ratio of this shape is locked.",
        kind: "Property",
        signature: "Excel.Shape.lockAspectRatio: boolean",
        examples: ["shape.lockAspectRatio = true;"],
      },
      {
        name: "Excel.Shape.name",
        description: "Specifies the name of the shape.",
        kind: "Property",
        signature: "Excel.Shape.name: string",
        examples: [
          'line.name = "StraightLine";',
          'shapeGroup.name = "Group";',
          'textbox.name = "Textbox";',
        ],
      },
      {
        name: "Excel.Shape.parentGroup",
        description: "Specifies the parent group of this shape.",
        kind: "Property",
        signature: "Excel.Shape.parentGroup: Shape",
        examples: [],
      },
      {
        name: "Excel.Shape.placement",
        description: "Represents how the object is attached to the cells below it.",
        kind: "Property",
        signature: 'Excel.Shape.placement: Placement | "TwoCell" | "OneCell" | "Absolute"',
        examples: [],
      },
      {
        name: "Excel.Shape.rotation",
        description: "Specifies the rotation, in degrees, of the shape.",
        kind: "Property",
        signature: "Excel.Shape.rotation: number",
        examples: ["shape.rotation = 45;"],
      },
      {
        name: "Excel.Shape.scriptLink",
        description:
          "Specifies the share link to an Office Script file on OneDrive that will be associated with this shape.",
        kind: "Property",
        signature: "Excel.Shape.scriptLink: string",
        examples: [],
      },
      {
        name: "Excel.Shape.textFrame",
        description: "Returns the text frame object of this shape.",
        kind: "Property",
        signature: "Excel.Shape.textFrame: Excel.TextFrame",
        examples: [
          "textbox.textFrame.autoSizeSetting = Excel.ShapeAutoSize.autoSizeShapeToFitText;",
          "textbox.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;",
          "textbox.textFrame.deleteText();",
        ],
      },
      {
        name: "Excel.Shape.top",
        description:
          "The distance, in points, from the top edge of the shape to the top edge of the worksheet. Throws an `InvalidArgument` exception when set with a negative value as an input.",
        kind: "Property",
        signature: "Excel.Shape.top: number",
        examples: ["shape.top = 5;", "shape.top = 100;", "shape.top = 300;", "textbox.top = 100;"],
      },
      {
        name: "Excel.Shape.type",
        description: "Returns the type of this shape. See `Excel.ShapeType` for details.",
        kind: "Property",
        signature:
          'Excel.Shape.type: "Unsupported" | "Line" | ShapeType | "Image" | "GeometricShape" | "Group"',
        examples: [],
      },
      {
        name: "Excel.Shape.visible",
        description: "Specifies if the shape is visible.",
        kind: "Property",
        signature: "Excel.Shape.visible: boolean",
        examples: [],
      },
      {
        name: "Excel.Shape.width",
        description:
          "Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value or zero as an input.",
        kind: "Property",
        signature: "Excel.Shape.width: number",
        examples: ["shape.width = 200;", "shape.width = 100;", "textbox.width = 175;"],
      },
      {
        name: "Excel.Shape.zOrderPosition",
        description:
          "Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.",
        kind: "Property",
        signature: "Excel.Shape.zOrderPosition: number",
        examples: [],
      },
      {
        name: "Excel.Shape.activate",
        description: "Activates the shape in the Excel UI.",
        kind: "Method",
        signature: "Excel.Shape.activate => () => void",
        examples: [],
      },
      {
        name: "Excel.Shape.copyTo",
        description:
          "Copies and pastes a `Shape` object. The pasted shape is copied to the same pixel location as this shape.",
        kind: "Method",
        signature: "Excel.Shape.copyTo => (destinationSheet?: Worksheet | string) => Excel.Shape",
        examples: [],
      },
      {
        name: "Excel.Shape.delete",
        description: "Removes the shape from the worksheet.",
        kind: "Method",
        signature: "Excel.Shape.delete() => void",
        examples: ["shapes.items.forEach((shape) => shape.delete());"],
      },
      {
        name: "Excel.Shape.getAsImage",
        description:
          "Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, and `Excel.PictureFormat.GIF`.",
        kind: "Method",
        signature:
          "Excel.Shape.getAsImage(format: Excel.PictureFormat): OfficeExtension.ClientResult<string>",
        examples: [
          "let stringResult = shape.getAsImage(Excel.PictureFormat.png);",
          "const result = shape.getAsImage(Excel.PictureFormat.png);",
        ],
      },
      {
        name: "Excel.Shape.incrementLeft",
        description: "Moves the shape horizontally by the specified number of points.",
        kind: "Method",
        signature: "Excel.Shape.incrementLeft(increment: number) => void",
        examples: ["shape.incrementLeft(-25);"],
      },
      {
        name: "Excel.Shape.incrementRotation",
        description:
          "Rotates the shape clockwise around the z-axis by the specified number of degrees. Use the `rotation` property to set the absolute rotation of the shape.",
        kind: "Method",
        signature: "Excel.Shape.incrementRotation(increment: number) => void",
        examples: ["shape.incrementRotation(180);"],
      },
      {
        name: "Excel.Shape.incrementTop",
        description: "Moves the shape vertically by the specified number of points.",
        kind: "Method",
        signature: "Excel.Shape.incrementTop(increment: number) => void",
        examples: ["shape.incrementTop(25);"],
      },
      {
        name: "Excel.Shape.scaleHeight",
        description:
          "Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.",
        kind: "Method",
        signature:
          "Excel.Shape.scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom): void",
        examples: ["shape.scaleHeight(1.25, Excel.ShapeScaleType.currentSize);"],
      },
      {
        name: "Excel.Shape.scaleWidth",
        description:
          "Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.",
        kind: "Method",
        signature:
          'Excel.Shape.scaleWidth => { (scaleFactor: number, scaleType: ShapeScaleType, scaleFrom?: ShapeScaleFrom): void; (scaleFactor: number, scaleType: "CurrentSize" | "OriginalSize", scaleFrom?: "ScaleFromTopLeft" | ... 1 more ... | "ScaleFromBottomRight"): void; (scaleFactor: number, scaleType: string, scaleFrom?: string): void; }',
        examples: [],
      },
      {
        name: "Excel.Shape.setZOrder",
        description:
          "Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.",
        kind: "Method",
        signature: "Excel.Shape.setZOrder(position: Excel.ShapeZOrder): void",
        examples: ["shape.setZOrder(Excel.ShapeZOrder.sendBackward);"],
      },
    ],
  },
  {
    objName: "Excel.ShapeCollection",
    apiList: [
      {
        name: "Excel.ShapeCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.ShapeCollection.items: Excel.Shape[]",
        examples: ["shapes.items.forEach((shape) => shape.delete());"],
      },
      {
        name: "Excel.ShapeCollection.addGeometricShape",
        description:
          "Adds a geometric shape to the worksheet. Returns a `Shape` object that represents the new shape.",
        kind: "Method",
        signature:
          "Excel.ShapeCollection.addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape",
        examples: [
          "const shape = activeWorksheet.shapes.addGeometricShape(Excel.GeometricShapeType.hexagon);",
          "const shape = activeWorksheet.shapes.addGeometricShape(Excel.GeometricShapeType.smileyFace);",
          "const shape = activeWorksheet.shapes.addGeometricShape(Excel.GeometricShapeType.triangle);",
        ],
      },
      {
        name: "Excel.ShapeCollection.addGroup",
        description:
          "Groups a subset of shapes in this collection's worksheet. Returns a `Shape` object that represents the new group of shapes.",
        kind: "Method",
        signature:
          "Excel.ShapeCollection.addGroup(values: (string | Excel.Shape)[]) => Excel.Shape",
        examples: [
          "const shapeGroup = activeWorksheet.shapes.addGroup([square, pentagon, octagon]);",
        ],
      },
      {
        name: "Excel.ShapeCollection.addImage",
        description:
          "Creates an image from a base64-encoded string and adds it to the worksheet. Returns the `Shape` object that represents the new image.",
        kind: "Method",
        signature: "Excel.ShapeCollection.addImage => (base64ImageString: string) => Excel.Shape",
        examples: [],
      },
      {
        name: "Excel.ShapeCollection.addLine",
        description:
          "Adds a line to worksheet. Returns a `Shape` object that represents the new line.",
        kind: "Method",
        signature:
          "Excel.ShapeCollection.addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape",
        examples: ["const line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);"],
      },
      {
        name: "Excel.ShapeCollection.addSvg",
        description:
          "Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet. Returns a `Shape` object that represents the new image.",
        kind: "Method",
        signature: "Excel.ShapeCollection.addSvg => (xml: string) => Excel.Shape",
        examples: [],
      },
      {
        name: "Excel.ShapeCollection.addTextBox",
        description:
          "Adds a text box to the worksheet with the provided text as the content. Returns a `Shape` object that represents the new text box.",
        kind: "Method",
        signature: "Excel.ShapeCollection.addTextBox(text?: string) => Excel.Shape",
        examples: ['const textbox = shapes.addTextBox("A box with text");'],
      },
      {
        name: "Excel.ShapeCollection.getCount",
        description: "Returns the number of shapes in the worksheet.",
        kind: "Method",
        signature: "Excel.ShapeCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.ShapeCollection.getItem",
        description: "Gets a shape using its name or ID.",
        kind: "Method",
        signature: "Excel.ShapeCollection.getItem(key: string) => Excel.Shape",
        examples: [
          'let shape = shapes.getItem("Image");',
          'const line = shapes.getItem("StraightLine").line;',
          'const image = activeWorksheet.shapes.getItem("Image").image;',
          'line.connectBeginShape(shapes.getItem("Left"), 2);',
          'line.connectEndShape(shapes.getItem("Right"), 0);',
          'const shape = activeWorksheet.shapes.getItem("Image");',
          'const shapeGroup = activeWorksheet.shapes.getItem("Group").group;',
          'const shape = activeWorksheet.shapes.getItem("Square");',
          'const shape = activeWorksheet.shapes.getItem("Pentagon");',
          'const shape = activeWorksheet.shapes.getItem("Octagon");',
          'const textbox = shapes.getItem("Textbox");',
          'const square = activeWorksheet.shapes.getItem("Square");',
          'const pentagon = activeWorksheet.shapes.getItem("Pentagon");',
          'const octagon = activeWorksheet.shapes.getItem("Octagon");',
        ],
      },
      {
        name: "Excel.ShapeCollection.getItemAt",
        description: "Gets a shape using its position in the collection.",
        kind: "Method",
        signature: "Excel.ShapeCollection.getItemAt => (index: number) => Excel.Shape",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ShapeFill",
    apiList: [
      {
        name: "Excel.ShapeFill.foregroundColor",
        description:
          'Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")',
        kind: "Property",
        signature: "Excel.ShapeFill.foregroundColor: string",
        examples: ['shape.fill.foregroundColor = "yellow";'],
      },
      {
        name: "Excel.ShapeFill.transparency",
        description:
          "Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.",
        kind: "Property",
        signature: "Excel.ShapeFill.transparency: number",
        examples: [],
      },
      {
        name: "Excel.ShapeFill.type",
        description: "Returns the fill type of the shape. See `Excel.ShapeFillType` for details.",
        kind: "Property",
        signature:
          'Excel.ShapeFill.type: "Solid" | ShapeFillType | "NoFill" | "Gradient" | "Pattern" | "PictureAndTexture" | "Mixed"',
        examples: [],
      },
      {
        name: "Excel.ShapeFill.clear",
        description: "Clears the fill formatting of this shape.",
        kind: "Method",
        signature: "Excel.ShapeFill.clear() => void",
        examples: ["shape.fill.clear();"],
      },
      {
        name: "Excel.ShapeFill.setSolidColor",
        description:
          'Sets the fill formatting of the shape to a uniform color. This changes the fill type to "Solid".',
        kind: "Method",
        signature: "Excel.ShapeFill.setSolidColor => (color: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ShapeFont",
    apiList: [
      {
        name: "Excel.ShapeFont.bold",
        description:
          "Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.",
        kind: "Property",
        signature: "Excel.ShapeFont.bold: boolean",
        examples: [],
      },
      {
        name: "Excel.ShapeFont.color",
        description:
          'HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.',
        kind: "Property",
        signature: "Excel.ShapeFont.color: string",
        examples: [],
      },
      {
        name: "Excel.ShapeFont.italic",
        description:
          "Represents the italic status of font. Returns `null` if the `TextRange` includes both italic and non-italic text fragments.",
        kind: "Property",
        signature: "Excel.ShapeFont.italic: boolean",
        examples: [],
      },
      {
        name: "Excel.ShapeFont.name",
        description:
          'Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.',
        kind: "Property",
        signature: "Excel.ShapeFont.name: string",
        examples: [],
      },
      {
        name: "Excel.ShapeFont.size",
        description:
          "Represents font size in points (e.g., 11). Returns `null` if the `TextRange` includes text fragments with different font sizes.",
        kind: "Property",
        signature: "Excel.ShapeFont.size: number",
        examples: [],
      },
      {
        name: "Excel.ShapeFont.underline",
        description:
          "Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See `Excel.ShapeFontUnderlineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ShapeFont.underline: "Double" | "None" | "Single" | "Dash" | ShapeFontUnderlineStyle | "Heavy" | "Dotted" | "DottedHeavy" | "DashHeavy" | "DashLong" | "DashLongHeavy" | ... 6 more ... | "WavyDouble"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ShapeGroup",
    apiList: [
      {
        name: "Excel.ShapeGroup.id",
        description: "Specifies the shape identifier.",
        kind: "Property",
        signature: "Excel.ShapeGroup.id: string",
        examples: [],
      },
      {
        name: "Excel.ShapeGroup.shape",
        description: "Returns the `Shape` object associated with the group.",
        kind: "Property",
        signature: "Excel.ShapeGroup.shape: Shape",
        examples: [],
      },
      {
        name: "Excel.ShapeGroup.shapes",
        description: "Returns the collection of `Shape` objects.",
        kind: "Property",
        signature: "Excel.ShapeGroup.shapes: GroupShapeCollection",
        examples: [],
      },
      {
        name: "Excel.ShapeGroup.ungroup",
        description: "Ungroups any grouped shapes in the specified shape group.",
        kind: "Method",
        signature: "Excel.ShapeGroup.ungroup() => void",
        examples: ["shapeGroup.ungroup();"],
      },
    ],
  },
  {
    objName: "Excel.ShapeLineFormat",
    apiList: [
      {
        name: "Excel.ShapeLineFormat.color",
        description:
          'Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").',
        kind: "Property",
        signature: "Excel.ShapeLineFormat.color: string",
        examples: [],
      },
      {
        name: "Excel.ShapeLineFormat.dashStyle",
        description:
          "Represents the line style of the shape. Returns `null` when the line is not visible or there are inconsistent dash styles. See `Excel.ShapeLineDashStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ShapeLineFormat.dashStyle: "Solid" | "Dash" | "DashDot" | "DashDotDot" | "RoundDot" | ShapeLineDashStyle | "LongDash" | "LongDashDot" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot"',
        examples: [],
      },
      {
        name: "Excel.ShapeLineFormat.style",
        description:
          "Represents the line style of the shape. Returns `null` when the line is not visible or there are inconsistent styles. See `Excel.ShapeLineStyle` for details.",
        kind: "Property",
        signature:
          'Excel.ShapeLineFormat.style: "Single" | ShapeLineStyle | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin"',
        examples: [],
      },
      {
        name: "Excel.ShapeLineFormat.transparency",
        description:
          "Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` when the shape has inconsistent transparencies.",
        kind: "Property",
        signature: "Excel.ShapeLineFormat.transparency: number",
        examples: [],
      },
      {
        name: "Excel.ShapeLineFormat.visible",
        description:
          "Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.",
        kind: "Property",
        signature: "Excel.ShapeLineFormat.visible: boolean",
        examples: [],
      },
      {
        name: "Excel.ShapeLineFormat.weight",
        description:
          "Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.",
        kind: "Property",
        signature: "Excel.ShapeLineFormat.weight: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ShowAsRule",
    apiList: [
      {
        name: "Excel.ShowAsRule.baseField",
        description:
          "The PivotField to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.",
        kind: "Property",
        signature: "Excel.ShowAsRule.baseField: Excel.PivotField",
        examples: [
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");',
          'wholesaleShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");',
        ],
      },
      {
        name: "Excel.ShowAsRule.baseItem",
        description:
          "The item to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.",
        kind: "Property",
        signature: "Excel.ShowAsRule.baseItem: Excel.PivotItem",
        examples: [
          'farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
          'wholesaleShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");',
        ],
      },
      {
        name: "Excel.ShowAsRule.calculation",
        description:
          "The `ShowAs` calculation to use for the PivotField. See `Excel.ShowAsCalculation` for details.",
        kind: "Property",
        signature:
          'Excel.ShowAsRule.calculation: Excel.ShowAsCalculation | "Unknown" | "None" | "PercentOfGrandTotal" | "PercentOfRowTotal" | "PercentOfColumnTotal" | "PercentOfParentRowTotal" | "PercentOfParentColumnTotal" | ... 8 more ... | "Index"',
        examples: [
          "farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;",
          "farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;",
          "wholesaleShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;",
          "wholesaleShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;",
        ],
      },
    ],
  },
  {
    objName: "Excel.Slicer",
    apiList: [
      {
        name: "Excel.Slicer.caption",
        description: "Represents the caption of the slicer.",
        kind: "Property",
        signature: "Excel.Slicer.caption: string",
        examples: ['slicer.caption = "Fruit Types";'],
      },
      {
        name: "Excel.Slicer.columnCount",
        description:
          "Represents the number of columns in the specified slicer. The default value is 1. Throws an `InvalidArgument` exception when set with a negative value or zero as an input.",
        kind: "Property",
        signature: "Excel.Slicer.columnCount: number",
        examples: [],
      },
      {
        name: "Excel.Slicer.disableMoveResizeUI",
        description:
          "Represents whether the specified slicer can be moved or resized. Value is `true` if the slicer cannot be moved or resized; otherwise `false`. The default value is `false`.",
        kind: "Property",
        signature: "Excel.Slicer.disableMoveResizeUI: boolean",
        examples: [],
      },
      {
        name: "Excel.Slicer.displayHeader",
        description:
          "Represents whether the header that displays the slicer caption is visible. Value is `true` if the header is visible; otherwise `false`. The default value is `true`.",
        kind: "Property",
        signature: "Excel.Slicer.displayHeader: boolean",
        examples: [],
      },
      {
        name: "Excel.Slicer.height",
        description:
          "Represents the height, in points, of the slicer. Throws an `InvalidArgument` exception when set with a negative value or zero as an input.",
        kind: "Property",
        signature: "Excel.Slicer.height: number",
        examples: ["slicer.height = 135;"],
      },
      {
        name: "Excel.Slicer.id",
        description: "Represents the unique ID of the slicer.",
        kind: "Property",
        signature: "Excel.Slicer.id: string",
        examples: [],
      },
      {
        name: "Excel.Slicer.isFilterCleared",
        description: "Value is `true` if all filters currently applied on the slicer are cleared.",
        kind: "Property",
        signature: "Excel.Slicer.isFilterCleared: boolean",
        examples: [],
      },
      {
        name: "Excel.Slicer.left",
        description:
          "Represents the distance, in points, from the left side of the slicer to the left of the worksheet. Throws an `InvalidArgument` error when set with a negative value as an input.",
        kind: "Property",
        signature: "Excel.Slicer.left: number",
        examples: ["slicer.left = 395;"],
      },
      {
        name: "Excel.Slicer.name",
        description: "Represents the name of the slicer.",
        kind: "Property",
        signature: "Excel.Slicer.name: string",
        examples: ['slicer.name = "Fruit Slicer";'],
      },
      {
        name: "Excel.Slicer.nameInFormula",
        description: "Represents the slicer name used in the formula.",
        kind: "Property",
        signature: "Excel.Slicer.nameInFormula: string",
        examples: [],
      },
      {
        name: "Excel.Slicer.rowHeight",
        description:
          "Represents the row height of the specified slicer. Throws an `InvalidArgument` exception when set with a negative value or zero as an input.",
        kind: "Property",
        signature: "Excel.Slicer.rowHeight: number",
        examples: [],
      },
      {
        name: "Excel.Slicer.slicerItems",
        description: "Represents the collection of slicer items that are part of the slicer.",
        kind: "Property",
        signature: "Excel.Slicer.slicerItems: SlicerItemCollection",
        examples: [],
      },
      {
        name: "Excel.Slicer.slicerStyle",
        description: "The style applied to the slicer.",
        kind: "Property",
        signature: "Excel.Slicer.slicerStyle: SlicerStyle",
        examples: [],
      },
      {
        name: "Excel.Slicer.sortBy",
        description:
          'Represents the sort order of the items in the slicer. Possible values are: "DataSourceOrder", "Ascending", "Descending".',
        kind: "Property",
        signature:
          'Excel.Slicer.sortBy: "Ascending" | "Descending" | SlicerSortType | "DataSourceOrder"',
        examples: [],
      },
      {
        name: "Excel.Slicer.sortUsingCustomLists",
        description:
          "Value is `true` if items in the specified slicer will be sorted by the custom lists.",
        kind: "Property",
        signature: "Excel.Slicer.sortUsingCustomLists: boolean",
        examples: [],
      },
      {
        name: "Excel.Slicer.style",
        description:
          'Constant value that represents the slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.',
        kind: "Property",
        signature: "Excel.Slicer.style: string",
        examples: ['slicer.style = "SlicerStyleLight6";'],
      },
      {
        name: "Excel.Slicer.top",
        description:
          "Represents the distance, in points, from the top edge of the slicer to the top of the worksheet. Throws an `InvalidArgument` error when set with a negative value as an input.",
        kind: "Property",
        signature: "Excel.Slicer.top: number",
        examples: ["slicer.top = 15;"],
      },
      {
        name: "Excel.Slicer.width",
        description:
          "Represents the width, in points, of the slicer. Throws an `InvalidArgument` error when set with a negative value or zero as an input.",
        kind: "Property",
        signature: "Excel.Slicer.width: number",
        examples: ["slicer.width = 150;"],
      },
      {
        name: "Excel.Slicer.worksheet",
        description: "Represents the worksheet containing the slicer.",
        kind: "Property",
        signature: "Excel.Slicer.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.Slicer.activate",
        description: "Activate the slicer in the Excel UI.",
        kind: "Method",
        signature: "Excel.Slicer.activate => () => void",
        examples: [],
      },
      {
        name: "Excel.Slicer.clearFilters",
        description: "Clears all the filters currently applied on the slicer.",
        kind: "Method",
        signature: "Excel.Slicer.clearFilters() => void",
        examples: ["slicer.clearFilters();"],
      },
      {
        name: "Excel.Slicer.delete",
        description: "Deletes the slicer.",
        kind: "Method",
        signature: "Excel.Slicer.delete() => void",
        examples: ["activeWorksheet.slicers.getItemAt(0).delete();"],
      },
      {
        name: "Excel.Slicer.getSelectedItems",
        description: "Returns an array of selected items' keys.",
        kind: "Method",
        signature: "Excel.Slicer.getSelectedItems => () => OfficeExtension.ClientResult<string[]>",
        examples: [],
      },
      {
        name: "Excel.Slicer.selectItems",
        description:
          "Selects slicer items based on their keys. The previous selections are cleared. All items will be selected by default if the array is empty.",
        kind: "Method",
        signature: "Excel.Slicer.selectItems(items?: string[]) => void",
        examples: ['slicer.selectItems(["Lemon", "Lime", "Orange"]);'],
      },
      {
        name: "Excel.Slicer.setStyle",
        description: "Sets the style applied to the slicer.",
        kind: "Method",
        signature:
          "Excel.Slicer.setStyle => (style: string | SlicerStyle | BuiltInSlicerStyle) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SlicerCollection",
    apiList: [
      {
        name: "Excel.SlicerCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.SlicerCollection.items: Slicer[]",
        examples: [],
      },
      {
        name: "Excel.SlicerCollection.add",
        description: "Adds a new slicer to the workbook.",
        kind: "Method",
        signature:
          "Excel.SlicerCollection.add(slicerSource: string | Excel.PivotTable | Excel.Table, sourceField: string | number | Excel.PivotField | Excel.TableColumn, slicerDestination?: string | Excel.Worksheet) => Excel.Slicer",
        examples: [
          'let slicer = activeWorksheet.slicers.add("Farm Sales", "Type");',
          'const slicer = activeWorksheet.slicers.add("Farm Sales", "Type");',
        ],
      },
      {
        name: "Excel.SlicerCollection.getCount",
        description: "Returns the number of slicers in the collection.",
        kind: "Method",
        signature: "Excel.SlicerCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.SlicerCollection.getItem",
        description: "Gets a slicer object using its name or ID.",
        kind: "Method",
        signature: "Excel.SlicerCollection.getItem(key: string) => Excel.Slicer",
        examples: [
          'let slicer = workbook.slicers.getItem("Fruit Slicer");',
          'const slicer = workbook.slicers.getItem("Fruit Slicer");',
        ],
      },
      {
        name: "Excel.SlicerCollection.getItemAt",
        description: "Gets a slicer based on its position in the collection.",
        kind: "Method",
        signature: "Excel.SlicerCollection.getItemAt(index: number) => Excel.Slicer",
        examples: ["activeWorksheet.slicers.getItemAt(0).delete();"],
      },
    ],
  },
  {
    objName: "Excel.SlicerItem",
    apiList: [
      {
        name: "Excel.SlicerItem.hasData",
        description: "Value is `true` if the slicer item has data.",
        kind: "Property",
        signature: "Excel.SlicerItem.hasData: boolean",
        examples: [],
      },
      {
        name: "Excel.SlicerItem.isSelected",
        description:
          "Value is `true` if the slicer item is selected. Setting this value will not clear the selected state of other slicer items. By default, if the slicer item is the only one selected, when it is deselected, all items will be selected.",
        kind: "Property",
        signature: "Excel.SlicerItem.isSelected: boolean",
        examples: [],
      },
      {
        name: "Excel.SlicerItem.key",
        description: "Represents the unique value representing the slicer item.",
        kind: "Property",
        signature: "Excel.SlicerItem.key: string",
        examples: [],
      },
      {
        name: "Excel.SlicerItem.name",
        description: "Represents the title displayed in the Excel UI.",
        kind: "Property",
        signature: "Excel.SlicerItem.name: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SlicerItemCollection",
    apiList: [
      {
        name: "Excel.SlicerItemCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.SlicerItemCollection.items: SlicerItem[]",
        examples: [],
      },
      {
        name: "Excel.SlicerItemCollection.getCount",
        description: "Returns the number of slicer items in the slicer.",
        kind: "Method",
        signature:
          "Excel.SlicerItemCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.SlicerItemCollection.getItem",
        description: "Gets a slicer item object using its key or name.",
        kind: "Method",
        signature: "Excel.SlicerItemCollection.getItem => (key: string) => Excel.SlicerItem",
        examples: [],
      },
      {
        name: "Excel.SlicerItemCollection.getItemAt",
        description: "Gets a slicer item based on its position in the collection.",
        kind: "Method",
        signature: "Excel.SlicerItemCollection.getItemAt => (index: number) => Excel.SlicerItem",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SlicerStyle",
    apiList: [
      {
        name: "Excel.SlicerStyle.name",
        description: "Specifies the name of the slicer style.",
        kind: "Property",
        signature: "Excel.SlicerStyle.name: string",
        examples: [],
      },
      {
        name: "Excel.SlicerStyle.readOnly",
        description: "Specifies if this `SlicerStyle` object is read-only.",
        kind: "Property",
        signature: "Excel.SlicerStyle.readOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.SlicerStyle.delete",
        description: "Deletes the slicer style.",
        kind: "Method",
        signature: "Excel.SlicerStyle.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.SlicerStyle.duplicate",
        description:
          "Creates a duplicate of this slicer style with copies of all the style elements.",
        kind: "Method",
        signature: "Excel.SlicerStyle.duplicate => () => Excel.SlicerStyle",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SlicerStyleCollection",
    apiList: [
      {
        name: "Excel.SlicerStyleCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.SlicerStyleCollection.items: SlicerStyle[]",
        examples: [],
      },
      {
        name: "Excel.SlicerStyleCollection.add",
        description: "Creates a blank slicer style with the specified name.",
        kind: "Method",
        signature:
          "Excel.SlicerStyleCollection.add => (name: string, makeUniqueName?: boolean) => Excel.SlicerStyle",
        examples: [],
      },
      {
        name: "Excel.SlicerStyleCollection.getCount",
        description: "Gets the number of slicer styles in the collection.",
        kind: "Method",
        signature:
          "Excel.SlicerStyleCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.SlicerStyleCollection.getDefault",
        description: "Gets the default `SlicerStyle` for the parent object's scope.",
        kind: "Method",
        signature: "Excel.SlicerStyleCollection.getDefault => () => Excel.SlicerStyle",
        examples: [],
      },
      {
        name: "Excel.SlicerStyleCollection.getItem",
        description: "Gets a `SlicerStyle` by name.",
        kind: "Method",
        signature: "Excel.SlicerStyleCollection.getItem => (name: string) => Excel.SlicerStyle",
        examples: [],
      },
      {
        name: "Excel.SlicerStyleCollection.setDefault",
        description: "Sets the default slicer style for use in the parent object's scope.",
        kind: "Method",
        signature:
          "Excel.SlicerStyleCollection.setDefault => (newDefaultStyle: SlicerStyle | string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.SpillErrorCellValue",
    apiList: [
      {
        name: "Excel.SpillErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.SpillErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.SpillErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.SpillErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.SpillErrorCellValue.columnCount",
        description:
          "Represents the number of columns that would spill if there were no #SPILL! error.",
        kind: "Property",
        signature: "Excel.SpillErrorCellValue.columnCount: number",
        examples: [],
      },
      {
        name: "Excel.SpillErrorCellValue.errorSubType",
        description: "Represents the type of `SpillErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.SpillErrorCellValue.errorSubType: "Unknown" | "Table" | SpillErrorCellValueSubType | "Collision" | "IndeterminateSize" | "WorksheetEdge" | "OutOfMemoryWhileCalc" | "MergedCell"',
        examples: [],
      },
      {
        name: "Excel.SpillErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.SpillErrorCellValue.errorType: ErrorCellValueType.spill | "Spill"',
        examples: [],
      },
      {
        name: "Excel.SpillErrorCellValue.rowCount",
        description:
          "Represents the number of rows that would spill if there were no #SPILL! error.",
        kind: "Property",
        signature: "Excel.SpillErrorCellValue.rowCount: number",
        examples: [],
      },
      {
        name: "Excel.SpillErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.SpillErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.StringCellValue",
    apiList: [
      {
        name: "Excel.StringCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.StringCellValue.basicType: RangeValueType.string | "String"',
        examples: [],
      },
      {
        name: "Excel.StringCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: "Excel.StringCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.StringCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.StringCellValue.type: CellValueType.string | "String"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Style",
    apiList: [
      {
        name: "Excel.Style.autoIndent",
        description:
          "Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.",
        kind: "Property",
        signature: "Excel.Style.autoIndent: boolean",
        examples: ["newStyle.autoIndent = true;"],
      },
      {
        name: "Excel.Style.borders",
        description:
          "A collection of four border objects that represent the style of the four borders.",
        kind: "Property",
        signature: "Excel.Style.borders: RangeBorderCollection",
        examples: [],
      },
      {
        name: "Excel.Style.builtIn",
        description: "Specifies if the style is a built-in style.",
        kind: "Property",
        signature: "Excel.Style.builtIn: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.fill",
        description: "The fill of the style.",
        kind: "Property",
        signature: "Excel.Style.fill: Excel.RangeFill",
        examples: [
          '[\n    "Bold: " + style.font.bold,\n    "Font color: " + style.font.color,\n    "Italic: " + style.font.italic,\n    "Name: " + style.font.name,\n    "Size: " + style.font.size,\n    "Fill color: " + style.fill.color,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.Style.font",
        description: "A `Font` object that represents the font of the style.",
        kind: "Property",
        signature: "Excel.Style.font: Excel.RangeFont",
        examples: [
          '[\n    "Bold: " + style.font.bold,\n    "Font color: " + style.font.color,\n    "Italic: " + style.font.italic,\n    "Name: " + style.font.name,\n    "Size: " + style.font.size,\n    "Fill color: " + style.fill.color,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.Style.formulaHidden",
        description: "Specifies if the formula will be hidden when the worksheet is protected.",
        kind: "Property",
        signature: "Excel.Style.formulaHidden: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.horizontalAlignment",
        description:
          "Represents the horizontal alignment for the style. See `Excel.HorizontalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.Style.horizontalAlignment: Excel.HorizontalAlignment | "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed"',
        examples: [
          '[\n    "Orientation: " + style.textOrientation,\n    "Horizontal alignment: " + style.horizontalAlignment,\n    "Add indent: " + style.autoIndent,\n    "Reading order: " + style.readingOrder,\n    "Wrap text: " + style.wrapText,\n    "Include protection: " + style.includeProtection,\n    "Shrink to fit: " + style.shrinkToFit,\n    "Style locked: " + style.locked,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.Style.includeAlignment",
        description:
          "Specifies if the style includes the auto indent, horizontal alignment, vertical alignment, wrap text, indent level, and text orientation properties.",
        kind: "Property",
        signature: "Excel.Style.includeAlignment: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.includeBorder",
        description:
          "Specifies if the style includes the color, color index, line style, and weight border properties.",
        kind: "Property",
        signature: "Excel.Style.includeBorder: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.includeFont",
        description:
          "Specifies if the style includes the background, bold, color, color index, font style, italic, name, size, strikethrough, subscript, superscript, and underline font properties.",
        kind: "Property",
        signature: "Excel.Style.includeFont: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.includeNumber",
        description: "Specifies if the style includes the number format property.",
        kind: "Property",
        signature: "Excel.Style.includeNumber: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.includePatterns",
        description:
          "Specifies if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.",
        kind: "Property",
        signature: "Excel.Style.includePatterns: boolean",
        examples: [],
      },
      {
        name: "Excel.Style.includeProtection",
        description:
          "Specifies if the style includes the formula hidden and locked protection properties.",
        kind: "Property",
        signature: "Excel.Style.includeProtection: boolean",
        examples: ["newStyle.includeProtection = true;"],
      },
      {
        name: "Excel.Style.indentLevel",
        description: "An integer from 0 to 250 that indicates the indent level for the style.",
        kind: "Property",
        signature: "Excel.Style.indentLevel: number",
        examples: [],
      },
      {
        name: "Excel.Style.locked",
        description: "Specifies if the object is locked when the worksheet is protected.",
        kind: "Property",
        signature: "Excel.Style.locked: boolean",
        examples: ["newStyle.locked = false;"],
      },
      {
        name: "Excel.Style.name",
        description: "The name of the style.",
        kind: "Property",
        signature: "Excel.Style.name: string",
        examples: [],
      },
      {
        name: "Excel.Style.numberFormat",
        description: "The format code of the number format for the style.",
        kind: "Property",
        signature: "Excel.Style.numberFormat: string",
        examples: [],
      },
      {
        name: "Excel.Style.numberFormatLocal",
        description: "The localized format code of the number format for the style.",
        kind: "Property",
        signature: "Excel.Style.numberFormatLocal: string",
        examples: [],
      },
      {
        name: "Excel.Style.readingOrder",
        description: "The reading order for the style.",
        kind: "Property",
        signature:
          'Excel.Style.readingOrder: Excel.ReadingOrder | "Context" | "LeftToRight" | "RightToLeft"',
        examples: [
          '[\n    "Orientation: " + style.textOrientation,\n    "Horizontal alignment: " + style.horizontalAlignment,\n    "Add indent: " + style.autoIndent,\n    "Reading order: " + style.readingOrder,\n    "Wrap text: " + style.wrapText,\n    "Include protection: " + style.includeProtection,\n    "Shrink to fit: " + style.shrinkToFit,\n    "Style locked: " + style.locked,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.Style.shrinkToFit",
        description:
          "Specifies if text automatically shrinks to fit in the available column width.",
        kind: "Property",
        signature: "Excel.Style.shrinkToFit: boolean",
        examples: ["newStyle.shrinkToFit = true;"],
      },
      {
        name: "Excel.Style.textOrientation",
        description: "The text orientation for the style.",
        kind: "Property",
        signature: "Excel.Style.textOrientation: number",
        examples: ["newStyle.textOrientation = 38;"],
      },
      {
        name: "Excel.Style.verticalAlignment",
        description:
          "Specifies the vertical alignment for the style. See `Excel.VerticalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.Style.verticalAlignment: "Center" | "Justify" | "Distributed" | VerticalAlignment | "Top" | "Bottom"',
        examples: [],
      },
      {
        name: "Excel.Style.wrapText",
        description: "Specifies if Excel wraps the text in the object.",
        kind: "Property",
        signature: "Excel.Style.wrapText: boolean",
        examples: [
          '[\n    "Orientation: " + style.textOrientation,\n    "Horizontal alignment: " + style.horizontalAlignment,\n    "Add indent: " + style.autoIndent,\n    "Reading order: " + style.readingOrder,\n    "Wrap text: " + style.wrapText,\n    "Include protection: " + style.includeProtection,\n    "Shrink to fit: " + style.shrinkToFit,\n    "Style locked: " + style.locked,\n  ].join("\\n");',
        ],
      },
      {
        name: "Excel.Style.delete",
        description: "Deletes this style.",
        kind: "Method",
        signature: "Excel.Style.delete() => void",
        examples: ["style.delete();"],
      },
    ],
  },
  {
    objName: "Excel.StyleCollection",
    apiList: [
      {
        name: "Excel.StyleCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.StyleCollection.items: Style[]",
        examples: [],
      },
      {
        name: "Excel.StyleCollection.add",
        description: "Adds a new style to the collection.",
        kind: "Method",
        signature: "Excel.StyleCollection.add(name: string) => void",
        examples: ['styles.add("Diagonal Orientation Style");'],
      },
      {
        name: "Excel.StyleCollection.getCount",
        description: "Gets the number of styles in the collection.",
        kind: "Method",
        signature: "Excel.StyleCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.StyleCollection.getItem",
        description: "Gets a `Style` by name.",
        kind: "Method",
        signature: "Excel.StyleCollection.getItem(name: string) => Excel.Style",
        examples: [
          'let style = workbook.styles.getItem("Diagonal Orientation Style");',
          'let style = workbook.styles.getItem("Normal");',
          'let newStyle = styles.getItem("Diagonal Orientation Style");',
        ],
      },
      {
        name: "Excel.StyleCollection.getItemAt",
        description: "Gets a style based on its position in the collection.",
        kind: "Method",
        signature: "Excel.StyleCollection.getItemAt => (index: number) => Excel.Style",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Subtotals",
    apiList: [
      {
        name: "Excel.Subtotals.automatic",
        description:
          "If `Automatic` is set to `true`, then all other values will be ignored when setting the `Subtotals`.",
        kind: "Property",
        signature: "Excel.Subtotals.automatic: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.average",
        kind: "Property",
        signature: "Excel.Subtotals.average: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.count",
        kind: "Property",
        signature: "Excel.Subtotals.count: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.countNumbers",
        kind: "Property",
        signature: "Excel.Subtotals.countNumbers: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.max",
        kind: "Property",
        signature: "Excel.Subtotals.max: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.min",
        kind: "Property",
        signature: "Excel.Subtotals.min: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.product",
        kind: "Property",
        signature: "Excel.Subtotals.product: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.standardDeviation",
        kind: "Property",
        signature: "Excel.Subtotals.standardDeviation: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.standardDeviationP",
        kind: "Property",
        signature: "Excel.Subtotals.standardDeviationP: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.sum",
        kind: "Property",
        signature: "Excel.Subtotals.sum: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.variance",
        kind: "Property",
        signature: "Excel.Subtotals.variance: boolean",
        examples: [],
      },
      {
        name: "Excel.Subtotals.varianceP",
        kind: "Property",
        signature: "Excel.Subtotals.varianceP: boolean",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Table",
    apiList: [
      {
        name: "Excel.Table.autoFilter",
        description: "Represents the `AutoFilter` object of the table.",
        kind: "Property",
        signature: "Excel.Table.autoFilter: Excel.AutoFilter",
        examples: [
          'activeTable.autoFilter.apply(activeTable.getRange(), 2, {\n    filterOn: Excel.FilterOn.values,\n    values: ["Restaurant", "Groceries"],\n  });',
          "activeTable.autoFilter.apply(activeTable.getRange(), 3, {\n    filterOn: Excel.FilterOn.dynamic,\n    dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage,\n  });",
        ],
      },
      {
        name: "Excel.Table.columns",
        description: "Represents a collection of all the columns in the table.",
        kind: "Property",
        signature: "Excel.Table.columns: Excel.TableColumnCollection",
        examples: [
          'let commentsRange = activeTable.columns.getItem("Comments").getDataBodyRange();',
          'activeTable.columns.items[0].name = "Purchase date";',
          'let columnRange = activeTable.columns.getItem("Merchant").getDataBodyRange().load("values");',
          'let categoryFilter = activeTable.columns.getItem("Category").filter;',
          'let amountFilter = activeTable.columns.getItem("Amount").filter;',
          'activeTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";',
          'const commentsRange = activeTable.columns.getItem("Comments").getDataBodyRange();',
          'const rankingRange = activeTable.columns.getItem("Ranking").getDataBodyRange();',
          'const nameRange = activeTable.columns.getItem("Baby Name").getDataBodyRange();',
          'let filter = activeTable.columns.getItem("Amount").filter;',
          'filter = activeTable.columns.getItem("Category").filter;',
          "const column = activeTable.columns.getItemAt(2);",
          "const column = activeTable.columns.getItemAt(0);",
          "const columns = activeTable.columns.getItemAt(0);",
          "const column = activeTable.columns.getItem(0);",
          "const column = activeTable.columns.add(null, values);",
          "const tableColumn = activeTable.columns.getItem(0);",
          'const salesColumn = activeTable.columns.getItem("Sales");',
          'const itemColumn = activeTable.columns.getItem("Item");',
          'const perYearColumns = activeTable.columns.items.filter((column) => column.name === "Per Year");',
          'const yearColumn = activeTable.columns.getItem("Year");',
          'const voltageColumn = activeTable.columns.getItem("Voltage");',
          'const reviewerColumn = activeTable.columns.getItem("Reviewer");',
          'const bookColumn = activeTable.columns.getItem("Book");',
          'const authorColumn = activeTable.columns.getItem("Author");',
          'const ratingColumn = activeTable.columns.getItem("Rating");',
        ],
      },
      {
        name: "Excel.Table.highlightFirstColumn",
        description: "Specifies if the first column contains special formatting.",
        kind: "Property",
        signature: "Excel.Table.highlightFirstColumn: boolean",
        examples: [],
      },
      {
        name: "Excel.Table.highlightLastColumn",
        description: "Specifies if the last column contains special formatting.",
        kind: "Property",
        signature: "Excel.Table.highlightLastColumn: boolean",
        examples: [],
      },
      {
        name: "Excel.Table.id",
        description:
          "Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.",
        kind: "Property",
        signature: "Excel.Table.id: string",
        examples: ["activeTable.id;"],
      },
      {
        name: "Excel.Table.legacyId",
        description: "Returns a numeric ID.",
        kind: "Property",
        signature: "Excel.Table.legacyId: string",
        examples: [],
      },
      {
        name: "Excel.Table.name",
        description:
          "Name of the table. The set name of the table must follow the guidelines specified in the Rename an Excel table article.",
        kind: "Property",
        signature: "Excel.Table.name: string",
        examples: [
          'expensesTable.name = "ExpensesTable";',
          'table.name = "Example";',
          "table.name;",
          'expensesTable.name = "SalesTable";',
          'activeTable.name = "Table1-Renamed";',
          "activeTable.name;",
          'newTable.name = "HighSalesLowRatings";',
        ],
      },
      {
        name: "Excel.Table.rows",
        description: "Represents a collection of all the rows in the table.",
        kind: "Property",
        signature: "Excel.Table.rows: Excel.TableRowCollection",
        examples: [
          'let rowRange = activeTable.rows.getItemAt(1).load("values");',
          'activeTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";',
          "expensesTable.rows.add(null, newData);",
          "const row = activeTable.rows.getItemAt(2);",
          "const row = activeTable.rows.getItemAt(0);",
          "const row = activeTable.rows.add(null, values);",
          "const tablerow = activeTable.rows.getItemAt(0);",
          "newTable.rows.add(null, newTableBody);",
        ],
      },
      {
        name: "Excel.Table.showBandedColumns",
        description:
          "Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.",
        kind: "Property",
        signature: "Excel.Table.showBandedColumns: boolean",
        examples: [],
      },
      {
        name: "Excel.Table.showBandedRows",
        description:
          "Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.",
        kind: "Property",
        signature: "Excel.Table.showBandedRows: boolean",
        examples: [],
      },
      {
        name: "Excel.Table.showFilterButton",
        description:
          "Specifies if the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.",
        kind: "Property",
        signature: "Excel.Table.showFilterButton: boolean",
        examples: [],
      },
      {
        name: "Excel.Table.showHeaders",
        description:
          "Specifies if the header row is visible. This value can be set to show or remove the header row.",
        kind: "Property",
        signature: "Excel.Table.showHeaders: boolean",
        examples: [],
      },
      {
        name: "Excel.Table.showTotals",
        description:
          "Specifies if the total row is visible. This value can be set to show or remove the total row.",
        kind: "Property",
        signature: "Excel.Table.showTotals: boolean",
        examples: ["activeTable.showTotals = false;"],
      },
      {
        name: "Excel.Table.sort",
        description: "Represents the sorting for the table.",
        kind: "Property",
        signature: "Excel.Table.sort: Excel.TableSort",
        examples: [
          "activeTable.sort.apply(\n    [\n      {\n        key: 2,\n        ascending: true,\n      },\n    ],\n    true\n  );",
        ],
      },
      {
        name: "Excel.Table.style",
        description:
          'Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.',
        kind: "Property",
        signature: "Excel.Table.style: string",
        examples: ['activeTable.style = "TableStyleMedium2";', "activeTable.style;"],
      },
      {
        name: "Excel.Table.tableStyle",
        description: "The style applied to the table.",
        kind: "Property",
        signature: "Excel.Table.tableStyle: TableStyle",
        examples: [],
      },
      {
        name: "Excel.Table.worksheet",
        description: "The worksheet containing the current table.",
        kind: "Property",
        signature: "Excel.Table.worksheet: Worksheet",
        examples: [],
      },
      {
        name: "Excel.Table.clearFilters",
        description: "Clears all the filters currently applied on the table.",
        kind: "Method",
        signature: "Excel.Table.clearFilters() => void",
        examples: ["activeTable.clearFilters();"],
      },
      {
        name: "Excel.Table.clearStyle",
        description: "Changes the table to use the default table style.",
        kind: "Method",
        signature: "Excel.Table.clearStyle => () => void",
        examples: [],
      },
      {
        name: "Excel.Table.convertToRange",
        description: "Converts the table into a normal range of cells. All data is preserved.",
        kind: "Method",
        signature: "Excel.Table.convertToRange() => Excel.Range",
        examples: ["activeTable.convertToRange();"],
      },
      {
        name: "Excel.Table.delete",
        description: "Deletes the table.",
        kind: "Method",
        signature: "Excel.Table.delete() => void",
        examples: ["activeTable.delete();"],
      },
      {
        name: "Excel.Table.getDataBodyRange",
        description: "Gets the range object associated with the data body of the table.",
        kind: "Method",
        signature: "Excel.Table.getDataBodyRange() => Excel.Range",
        examples: [
          "const temperatureDataRange = activeTable.getDataBodyRange();",
          'let bodyRange = activeTable.getDataBodyRange().load("values");',
          "let sortRange = activeTable.getDataBodyRange();",
          "let visibleRange = activeTable.getDataBodyRange().getVisibleView();",
          'activeTable.getDataBodyRange().format.fill.color = "#DAF7A6";',
          'table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];',
          'table.getDataBodyRange().getRow(1).values = [["D", 4]];',
          "let dataRange = activeTable.getDataBodyRange();",
          "const dataRange = activeTable.getDataBodyRange();",
          "const tableDataRange = activeTable.getDataBodyRange();",
          "const conditionalFormat = activeTable.getDataBodyRange().conditionalFormats.add(Excel.ConditionalFormatType.custom);",
          "const tableDataBody = activeTable.getDataBodyRange().values;",
        ],
      },
      {
        name: "Excel.Table.getHeaderRowRange",
        description: "Gets the range object associated with the header row of the table.",
        kind: "Method",
        signature: "Excel.Table.getHeaderRowRange() => Excel.Range",
        examples: [
          'expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];',
          'let headerRange = activeTable.getHeaderRowRange().load("values");',
          'activeTable.getHeaderRowRange().format.fill.color = "#C70039";',
          'expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];',
          "const tableHeaderRange = activeTable.getHeaderRowRange();",
          "newTable.getHeaderRowRange().values = activeTable.getHeaderRowRange().values;",
          "newTable.getHeaderRowRange().values = selectedRange.getRow(0).values;",
        ],
      },
      {
        name: "Excel.Table.getRange",
        description: "Gets the range object associated with the entire table.",
        kind: "Method",
        signature: "Excel.Table.getRange() => Excel.Range",
        examples: [
          "const activeTableRange = activeTable.getRange();",
          "activeTable.getRange().format.autofitColumns();",
          "const expensesTableValues = activeTable.getRange().values;",
        ],
      },
      {
        name: "Excel.Table.getTotalRowRange",
        description: "Gets the range object associated with the totals row of the table.",
        kind: "Method",
        signature: "Excel.Table.getTotalRowRange() => Excel.Range",
        examples: ["const tableTotalsRange = activeTable.getTotalRowRange();"],
      },
      {
        name: "Excel.Table.reapplyFilters",
        description: "Reapplies all the filters currently on the table.",
        kind: "Method",
        signature: "Excel.Table.reapplyFilters => () => void",
        examples: [],
      },
      {
        name: "Excel.Table.resize",
        description:
          "Resize the table to the new range. The new range must overlap with the original table range and the headers (or the top of the table) must be in the same row.",
        kind: "Method",
        signature: "Excel.Table.resize(newRange: string | Excel.Range) => void",
        examples: ['activeTable.resize("A1:D20");'],
      },
      {
        name: "Excel.Table.setStyle",
        description: "Sets the style applied to the table.",
        kind: "Method",
        signature:
          "Excel.Table.setStyle => (style: string | TableStyle | BuiltInTableStyle) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TableCollection",
    apiList: [
      {
        name: "Excel.TableCollection.count",
        description: "Returns the number of tables in the workbook.",
        kind: "Property",
        signature: "Excel.TableCollection.count: number",
        examples: ["tables.count;"],
      },
      {
        name: "Excel.TableCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.TableCollection.items: Table[]",
        examples: [],
      },
      {
        name: "Excel.TableCollection.add",
        description:
          "Creates a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.",
        kind: "Method",
        signature:
          "Excel.TableCollection.add(address: string | Excel.Range, hasHeaders: boolean) => Excel.Table",
        examples: [
          'activeWorksheet.tables.add("B2:E5", true);',
          'let expensesTable = activeWorksheet.tables.add("A1:D1", true);',
          'let expensesTable = activeWorksheet.tables.add("A1:E7", true);',
          'let table = activeWorksheet.tables.add("A1:B3", true);',
          'let expensesTable = sheet.tables.add("A1:E1", true);',
          'const table = workbook.tables.add("Sheet1!A1:E7", true);',
          'const newTable = activeWorksheet.tables.add("G1:K1", true);',
          'const newTable = activeWorksheet.tables.add("G1:J1", true);',
        ],
      },
      {
        name: "Excel.TableCollection.getCount",
        description: "Gets the number of tables in the collection.",
        kind: "Method",
        signature: "Excel.TableCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.TableCollection.getItem",
        description: "Gets a table by name or ID.",
        kind: "Method",
        signature: "Excel.TableCollection.getItem(key: string) => Excel.Table",
        examples: [
          'const activeTable = activeWorksheet.tables.getItem("TemperatureTable");',
          'const activeTable = activeWorksheet.tables.getItem("AthletesTable");',
          'const activeTable = activeWorksheet.tables.getItem("ExpensesTable");',
          'const activeTable = activeWorksheet.tables.getItem("SalesTable");',
          'const activeTable = activeWorksheet.tables.getItem("Sales");',
          'const activeTable = activeWorksheet.tables.getItem("Table1");',
          'const activeTable = activeWorksheet.tables.getItem("NameOptionsTable");',
          'const activeTable = activeWorksheet.tables.getItem("Table2");',
          'const activeTable = activeWorksheet.tables.getItem("Table5");',
          'const activeTable = activeWorksheet.tables.getItem("ProductSales");',
          'const activeTable = activeWorksheet.tables.getItem("UnfilteredTable");',
        ],
      },
      {
        name: "Excel.TableCollection.getItemAt",
        description: "Gets a table based on its position in the collection.",
        kind: "Method",
        signature: "Excel.TableCollection.getItemAt(index: number) => Excel.Table",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TableColumn",
    apiList: [
      {
        name: "Excel.TableColumn.filter",
        description: "Retrieves the filter applied to the column.",
        kind: "Property",
        signature: "Excel.TableColumn.filter: Excel.Filter",
        examples: [
          'let categoryFilter = activeTable.columns.getItem("Category").filter;',
          'let amountFilter = activeTable.columns.getItem("Amount").filter;',
          'let filter = activeTable.columns.getItem("Amount").filter;',
          'filter = activeTable.columns.getItem("Category").filter;',
        ],
      },
      {
        name: "Excel.TableColumn.id",
        description: "Returns a unique key that identifies the column within the table.",
        kind: "Property",
        signature: "Excel.TableColumn.id: number",
        examples: [],
      },
      {
        name: "Excel.TableColumn.index",
        description:
          "Returns the index number of the column within the columns collection of the table. Zero-indexed.",
        kind: "Property",
        signature: "Excel.TableColumn.index: number",
        examples: ["column.index;"],
      },
      {
        name: "Excel.TableColumn.name",
        description: "Specifies the name of the table column.",
        kind: "Property",
        signature: "Excel.TableColumn.name: string",
        examples: [
          'activeTable.columns.items[0].name = "Purchase date";',
          "column.name;",
          "tableColumn.name;",
          'const perYearColumns = activeTable.columns.items.filter((column) => column.name === "Per Year");',
        ],
      },
      {
        name: "Excel.TableColumn.values",
        description:
          'Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string. If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.',
        kind: "Property",
        signature: "Excel.TableColumn.values: any[][]",
        examples: [],
      },
      {
        name: "Excel.TableColumn.delete",
        description: "Deletes the column from the table.",
        kind: "Method",
        signature: "Excel.TableColumn.delete() => void",
        examples: ["column.delete();", "perYearColumns.forEach((column) => column.delete());"],
      },
      {
        name: "Excel.TableColumn.getDataBodyRange",
        description: "Gets the range object associated with the data body of the column.",
        kind: "Method",
        signature: "Excel.TableColumn.getDataBodyRange() => Excel.Range",
        examples: [
          'let commentsRange = activeTable.columns.getItem("Comments").getDataBodyRange();',
          'let columnRange = activeTable.columns.getItem("Merchant").getDataBodyRange().load("values");',
          'activeTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";',
          'const commentsRange = activeTable.columns.getItem("Comments").getDataBodyRange();',
          'const rankingRange = activeTable.columns.getItem("Ranking").getDataBodyRange();',
          'const nameRange = activeTable.columns.getItem("Baby Name").getDataBodyRange();',
          "const dataBodyRange = column.getDataBodyRange();",
          "const salesColumnValues = salesColumn.getDataBodyRange().values;",
          "const itemColumnValues = itemColumn.getDataBodyRange().values;",
          "salesColumn.getDataBodyRange().values = salesColumnValues;",
          "const yearColumnValues = yearColumn.getDataBodyRange().values;",
          "const voltageColumnValues = voltageColumn.getDataBodyRange().values;",
          "const reviewerColumnValues = reviewerColumn.getDataBodyRange().values;",
          "const bookColumnValues = bookColumn.getDataBodyRange().values;",
          "const authorColumnValues = authorColumn.getDataBodyRange().values;",
          "const ratingColumnValues = ratingColumn.getDataBodyRange().values;",
        ],
      },
      {
        name: "Excel.TableColumn.getHeaderRowRange",
        description: "Gets the range object associated with the header row of the column.",
        kind: "Method",
        signature: "Excel.TableColumn.getHeaderRowRange() => Excel.Range",
        examples: ["const headerRowRange = columns.getHeaderRowRange();"],
      },
      {
        name: "Excel.TableColumn.getRange",
        description: "Gets the range object associated with the entire column.",
        kind: "Method",
        signature: "Excel.TableColumn.getRange() => Excel.Range",
        examples: ["const columnRange = columns.getRange();"],
      },
      {
        name: "Excel.TableColumn.getTotalRowRange",
        description: "Gets the range object associated with the totals row of the column.",
        kind: "Method",
        signature: "Excel.TableColumn.getTotalRowRange() => Excel.Range",
        examples: ["const totalRowRange = columns.getTotalRowRange();"],
      },
    ],
  },
  {
    objName: "Excel.TableColumnCollection",
    apiList: [
      {
        name: "Excel.TableColumnCollection.count",
        description: "Returns the number of columns in the table.",
        kind: "Property",
        signature: "Excel.TableColumnCollection.count: number",
        examples: [],
      },
      {
        name: "Excel.TableColumnCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.TableColumnCollection.items: Excel.TableColumn[]",
        examples: [
          'activeTable.columns.items[0].name = "Purchase date";',
          'const perYearColumns = activeTable.columns.items.filter((column) => column.name === "Per Year");',
        ],
      },
      {
        name: "Excel.TableColumnCollection.add",
        description: "Adds a new column to the table.",
        kind: "Method",
        signature:
          "Excel.TableColumnCollection.add(index?: number, values?: string | number | boolean | (string | number | boolean)[][], name?: string) => Excel.TableColumn",
        examples: ["const column = activeTable.columns.add(null, values);"],
      },
      {
        name: "Excel.TableColumnCollection.addAsJson",
        description:
          "Adds a new column to the table. Unlike `add()`, `addAsJson()` takes any type of cell value, such as image or entity data types.",
        kind: "Method",
        signature:
          "Excel.TableColumnCollection.addAsJson => (index?: number, values?: CellValue[][], name?: string) => Excel.TableColumn",
        examples: [],
      },
      {
        name: "Excel.TableColumnCollection.getCount",
        description: "Gets the number of columns in the table.",
        kind: "Method",
        signature:
          "Excel.TableColumnCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.TableColumnCollection.getItem",
        description: "Gets a column object by name or ID.",
        kind: "Method",
        signature: "Excel.TableColumnCollection.getItem(key: string | number) => Excel.TableColumn",
        examples: [
          'let commentsRange = activeTable.columns.getItem("Comments").getDataBodyRange();',
          'let columnRange = activeTable.columns.getItem("Merchant").getDataBodyRange().load("values");',
          'let categoryFilter = activeTable.columns.getItem("Category").filter;',
          'let amountFilter = activeTable.columns.getItem("Amount").filter;',
          'const commentsRange = activeTable.columns.getItem("Comments").getDataBodyRange();',
          'const rankingRange = activeTable.columns.getItem("Ranking").getDataBodyRange();',
          'const nameRange = activeTable.columns.getItem("Baby Name").getDataBodyRange();',
          'let filter = activeTable.columns.getItem("Amount").filter;',
          'filter = activeTable.columns.getItem("Category").filter;',
          "const column = activeTable.columns.getItem(0);",
          "const tableColumn = activeTable.columns.getItem(0);",
          'const salesColumn = activeTable.columns.getItem("Sales");',
          'const itemColumn = activeTable.columns.getItem("Item");',
          'const yearColumn = activeTable.columns.getItem("Year");',
          'const voltageColumn = activeTable.columns.getItem("Voltage");',
          'const reviewerColumn = activeTable.columns.getItem("Reviewer");',
          'const bookColumn = activeTable.columns.getItem("Book");',
          'const authorColumn = activeTable.columns.getItem("Author");',
          'const ratingColumn = activeTable.columns.getItem("Rating");',
        ],
      },
      {
        name: "Excel.TableColumnCollection.getItemAt",
        description: "Gets a column based on its position in the collection.",
        kind: "Method",
        signature: "Excel.TableColumnCollection.getItemAt(index: number) => Excel.TableColumn",
        examples: [
          'activeTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";',
          "const column = activeTable.columns.getItemAt(2);",
          "const column = activeTable.columns.getItemAt(0);",
          "const columns = activeTable.columns.getItemAt(0);",
        ],
      },
    ],
  },
  {
    objName: "Excel.TableRow",
    apiList: [
      {
        name: "Excel.TableRow.index",
        description:
          "Returns the index number of the row within the rows collection of the table. Zero-indexed.",
        kind: "Property",
        signature: "Excel.TableRow.index: number",
        examples: ["row.index;"],
      },
      {
        name: "Excel.TableRow.values",
        description:
          'Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string. If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.',
        kind: "Property",
        signature: "Excel.TableRow.values: any[][]",
        examples: ["let secondRowValues = rowRange.values;", "tablerow.values;"],
      },
      {
        name: "Excel.TableRow.delete",
        description: "Deletes the row from the table.",
        kind: "Method",
        signature: "Excel.TableRow.delete() => void",
        examples: ["row.delete();"],
      },
      {
        name: "Excel.TableRow.getRange",
        description: "Returns the range object associated with the entire row.",
        kind: "Method",
        signature: "Excel.TableRow.getRange() => Excel.Range",
        examples: [
          'activeTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";',
          "const rowRange = row.getRange();",
        ],
      },
    ],
  },
  {
    objName: "Excel.TableRowCollection",
    apiList: [
      {
        name: "Excel.TableRowCollection.count",
        description: "Returns the number of rows in the table.",
        kind: "Property",
        signature: "Excel.TableRowCollection.count: number",
        examples: [],
      },
      {
        name: "Excel.TableRowCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.TableRowCollection.items: TableRow[]",
        examples: [],
      },
      {
        name: "Excel.TableRowCollection.add",
        description:
          "Adds one or more rows to the table. The return object will be the top of the newly added row(s). Note that unlike ranges or columns, which will adjust if new rows or columns are added before them, a `TableRow` object represents the physical location of the table row, but not the data. That is, if the data is sorted or if new rows are added, a table row will continue to point at the index for which it was created.",
        kind: "Method",
        signature:
          "Excel.TableRowCollection.add(index?: number, values?: string | number | boolean | (string | number | boolean)[][], alwaysInsert?: boolean) => Excel.TableRow",
        examples: [
          "expensesTable.rows.add(null, newData);",
          "const row = activeTable.rows.add(null, values);",
          "newTable.rows.add(null, newTableBody);",
        ],
      },
      {
        name: "Excel.TableRowCollection.addAsJson",
        description:
          "Adds one or more rows to the table. The returned object will be the top row of the newly added row or rows. Unlike `add()`, `addAsJson()` takes any type of cell value, such as image or entity data types. Note that unlike ranges or columns, which will adjust if new rows or columns are added before them, a `TableRow` object represents the physical location of the table row, but not the data. That is, if the data is sorted or if new rows are added, a table row will continue to point at the index for which it was created.",
        kind: "Method",
        signature:
          "Excel.TableRowCollection.addAsJson => (index?: number, values?: CellValue[][], alwaysInsert?: boolean) => Excel.TableRow",
        examples: [],
      },
      {
        name: "Excel.TableRowCollection.deleteRows",
        description:
          "Delete multiple rows from a table. These rows don't need to be sequential. This method will throw the `InvalidArgument` error if a chosen row has already been deleted or doesn't exist. This method will throw the `InsertDeleteConflict` error if the table on which the method is called has a filter applied.",
        kind: "Method",
        signature: "Excel.TableRowCollection.deleteRows => (rows: number[] | TableRow[]) => void",
        examples: [],
      },
      {
        name: "Excel.TableRowCollection.deleteRowsAt",
        description:
          "Delete a specified number of rows from a table, starting at a given index. This method will throw the `InsertDeleteConflict` error if the table on which the method is called has a filter applied.",
        kind: "Method",
        signature:
          "Excel.TableRowCollection.deleteRowsAt => (index: number, count?: number) => void",
        examples: [],
      },
      {
        name: "Excel.TableRowCollection.getCount",
        description: "Gets the number of rows in the table.",
        kind: "Method",
        signature:
          "Excel.TableRowCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.TableRowCollection.getItemAt",
        description:
          "Gets a row based on its position in the collection. Note that unlike ranges or columns, which will adjust if new rows or columns are added before them, a `TableRow` object represents the physical location of the table row, but not the data. That is, if the data is sorted or if new rows are added, a table row will continue to point at the index for which it was created.",
        kind: "Method",
        signature: "Excel.TableRowCollection.getItemAt(index: number) => Excel.TableRow",
        examples: [
          'let rowRange = activeTable.rows.getItemAt(1).load("values");',
          'activeTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";',
          "const row = activeTable.rows.getItemAt(2);",
          "const row = activeTable.rows.getItemAt(0);",
          "const tablerow = activeTable.rows.getItemAt(0);",
        ],
      },
    ],
  },
  {
    objName: "Excel.TableScopedCollection",
    apiList: [
      {
        name: "Excel.TableScopedCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.TableScopedCollection.items: Table[]",
        examples: [],
      },
      {
        name: "Excel.TableScopedCollection.getCount",
        description: "Gets the number of tables in the collection.",
        kind: "Method",
        signature:
          "Excel.TableScopedCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.TableScopedCollection.getFirst",
        description:
          "Gets the first table in the collection. The tables in the collection are sorted top-to-bottom and left-to-right, such that top-left table is the first table in the collection.",
        kind: "Method",
        signature: "Excel.TableScopedCollection.getFirst => () => Excel.Table",
        examples: [],
      },
      {
        name: "Excel.TableScopedCollection.getItem",
        description: "Gets a table by name or ID.",
        kind: "Method",
        signature: "Excel.TableScopedCollection.getItem => (key: string) => Excel.Table",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TableSort",
    apiList: [
      {
        name: "Excel.TableSort.fields",
        description: "Specifies the current conditions used to last sort the table.",
        kind: "Property",
        signature: "Excel.TableSort.fields: SortField[]",
        examples: [],
      },
      {
        name: "Excel.TableSort.matchCase",
        description: "Specifies if the casing impacts the last sort of the table.",
        kind: "Property",
        signature: "Excel.TableSort.matchCase: boolean",
        examples: [],
      },
      {
        name: "Excel.TableSort.method",
        description:
          "Represents the Chinese character ordering method last used to sort the table.",
        kind: "Property",
        signature: 'Excel.TableSort.method: SortMethod | "PinYin" | "StrokeCount"',
        examples: [],
      },
      {
        name: "Excel.TableSort.apply",
        description: "Perform a sort operation.",
        kind: "Method",
        signature:
          "Excel.TableSort.apply(fields: Excel.SortField[], matchCase?: boolean, method?: Excel.SortMethod): void",
        examples: [
          "activeTable.sort.apply(\n    [\n      {\n        key: 2,\n        ascending: true,\n      },\n    ],\n    true\n  );",
        ],
      },
      {
        name: "Excel.TableSort.clear",
        description:
          "Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.",
        kind: "Method",
        signature: "Excel.TableSort.clear => () => void",
        examples: [],
      },
      {
        name: "Excel.TableSort.reapply",
        description: "Reapplies the current sorting parameters to the table.",
        kind: "Method",
        signature: "Excel.TableSort.reapply => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TableStyle",
    apiList: [
      {
        name: "Excel.TableStyle.name",
        description: "Specifies the name of the table style.",
        kind: "Property",
        signature: "Excel.TableStyle.name: string",
        examples: [],
      },
      {
        name: "Excel.TableStyle.readOnly",
        description: "Specifies if this `TableStyle` object is read-only.",
        kind: "Property",
        signature: "Excel.TableStyle.readOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.TableStyle.delete",
        description: "Deletes the table style.",
        kind: "Method",
        signature: "Excel.TableStyle.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.TableStyle.duplicate",
        description:
          "Creates a duplicate of this table style with copies of all the style elements.",
        kind: "Method",
        signature: "Excel.TableStyle.duplicate => () => Excel.TableStyle",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TableStyleCollection",
    apiList: [
      {
        name: "Excel.TableStyleCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.TableStyleCollection.items: TableStyle[]",
        examples: [],
      },
      {
        name: "Excel.TableStyleCollection.add",
        description: "Creates a blank `TableStyle` with the specified name.",
        kind: "Method",
        signature:
          "Excel.TableStyleCollection.add => (name: string, makeUniqueName?: boolean) => Excel.TableStyle",
        examples: [],
      },
      {
        name: "Excel.TableStyleCollection.getCount",
        description: "Gets the number of table styles in the collection.",
        kind: "Method",
        signature:
          "Excel.TableStyleCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.TableStyleCollection.getDefault",
        description: "Gets the default table style for the parent object's scope.",
        kind: "Method",
        signature: "Excel.TableStyleCollection.getDefault => () => Excel.TableStyle",
        examples: [],
      },
      {
        name: "Excel.TableStyleCollection.getItem",
        description: "Gets a `TableStyle` by name.",
        kind: "Method",
        signature: "Excel.TableStyleCollection.getItem => (name: string) => Excel.TableStyle",
        examples: [],
      },
      {
        name: "Excel.TableStyleCollection.setDefault",
        description: "Sets the default table style for use in the parent object's scope.",
        kind: "Method",
        signature:
          "Excel.TableStyleCollection.setDefault => (newDefaultStyle: TableStyle | string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TextConditionalFormat",
    apiList: [
      {
        name: "Excel.TextConditionalFormat.format",
        description:
          "Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.",
        kind: "Property",
        signature: "Excel.TextConditionalFormat.format: Excel.ConditionalRangeFormat",
        examples: ['conditionalFormat.textComparison.format.font.color = "red";'],
      },
      {
        name: "Excel.TextConditionalFormat.rule",
        description: "The rule of the conditional format.",
        kind: "Property",
        signature: "Excel.TextConditionalFormat.rule: Excel.ConditionalTextComparisonRule",
        examples: [
          'conditionalFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Delayed" };',
        ],
      },
    ],
  },
  {
    objName: "Excel.TextFrame",
    apiList: [
      {
        name: "Excel.TextFrame.autoSizeSetting",
        description:
          "The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.",
        kind: "Property",
        signature:
          'Excel.TextFrame.autoSizeSetting: Excel.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed"',
        examples: [
          "textbox.textFrame.autoSizeSetting = Excel.ShapeAutoSize.autoSizeShapeToFitText;",
        ],
      },
      {
        name: "Excel.TextFrame.bottomMargin",
        description: "Represents the bottom margin, in points, of the text frame.",
        kind: "Property",
        signature: "Excel.TextFrame.bottomMargin: number",
        examples: [],
      },
      {
        name: "Excel.TextFrame.hasText",
        description: "Specifies if the text frame contains text.",
        kind: "Property",
        signature: "Excel.TextFrame.hasText: boolean",
        examples: [],
      },
      {
        name: "Excel.TextFrame.horizontalAlignment",
        description:
          "Represents the horizontal alignment of the text frame. See `Excel.ShapeTextHorizontalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.TextFrame.horizontalAlignment: Excel.ShapeTextHorizontalAlignment | "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed"',
        examples: [
          "textbox.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;",
        ],
      },
      {
        name: "Excel.TextFrame.horizontalOverflow",
        description:
          "Represents the horizontal overflow behavior of the text frame. See `Excel.ShapeTextHorizontalOverflow` for details.",
        kind: "Property",
        signature:
          'Excel.TextFrame.horizontalOverflow: ShapeTextHorizontalOverflow | "Overflow" | "Clip"',
        examples: [],
      },
      {
        name: "Excel.TextFrame.leftMargin",
        description: "Represents the left margin, in points, of the text frame.",
        kind: "Property",
        signature: "Excel.TextFrame.leftMargin: number",
        examples: [],
      },
      {
        name: "Excel.TextFrame.orientation",
        description:
          "Represents the angle to which the text is oriented for the text frame. See `Excel.ShapeTextOrientation` for details.",
        kind: "Property",
        signature:
          'Excel.TextFrame.orientation: "Horizontal" | "Vertical" | ShapeTextOrientation | "Vertical270" | "WordArtVertical" | "EastAsianVertical" | "MongolianVertical" | "WordArtVerticalRTL"',
        examples: [],
      },
      {
        name: "Excel.TextFrame.readingOrder",
        description:
          "Represents the reading order of the text frame, either left-to-right or right-to-left. See `Excel.ShapeTextReadingOrder` for details.",
        kind: "Property",
        signature:
          'Excel.TextFrame.readingOrder: "LeftToRight" | "RightToLeft" | ShapeTextReadingOrder',
        examples: [],
      },
      {
        name: "Excel.TextFrame.rightMargin",
        description: "Represents the right margin, in points, of the text frame.",
        kind: "Property",
        signature: "Excel.TextFrame.rightMargin: number",
        examples: [],
      },
      {
        name: "Excel.TextFrame.textRange",
        description:
          "Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See `Excel.TextRange` for details.",
        kind: "Property",
        signature: "Excel.TextFrame.textRange: TextRange",
        examples: [],
      },
      {
        name: "Excel.TextFrame.topMargin",
        description: "Represents the top margin, in points, of the text frame.",
        kind: "Property",
        signature: "Excel.TextFrame.topMargin: number",
        examples: [],
      },
      {
        name: "Excel.TextFrame.verticalAlignment",
        description:
          "Represents the vertical alignment of the text frame. See `Excel.ShapeTextVerticalAlignment` for details.",
        kind: "Property",
        signature:
          'Excel.TextFrame.verticalAlignment: "Distributed" | "Top" | "Bottom" | ShapeTextVerticalAlignment | "Middle" | "Justified"',
        examples: [],
      },
      {
        name: "Excel.TextFrame.verticalOverflow",
        description:
          "Represents the vertical overflow behavior of the text frame. See `Excel.ShapeTextVerticalOverflow` for details.",
        kind: "Property",
        signature:
          'Excel.TextFrame.verticalOverflow: "Overflow" | "Clip" | ShapeTextVerticalOverflow | "Ellipsis"',
        examples: [],
      },
      {
        name: "Excel.TextFrame.deleteText",
        description: "Deletes all the text in the text frame.",
        kind: "Method",
        signature: "Excel.TextFrame.deleteText() => void",
        examples: ["textbox.textFrame.deleteText();"],
      },
    ],
  },
  {
    objName: "Excel.TextRange",
    apiList: [
      {
        name: "Excel.TextRange.font",
        description:
          "Returns a `ShapeFont` object that represents the font attributes for the text range.",
        kind: "Property",
        signature: "Excel.TextRange.font: ShapeFont",
        examples: [],
      },
      {
        name: "Excel.TextRange.text",
        description: "Represents the plain text content of the text range.",
        kind: "Property",
        signature: "Excel.TextRange.text: string",
        examples: [],
      },
      {
        name: "Excel.TextRange.getSubstring",
        description: "Returns a TextRange object for the substring in the given range.",
        kind: "Method",
        signature:
          "Excel.TextRange.getSubstring => (start: number, length?: number) => Excel.TextRange",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TimelineStyle",
    apiList: [
      {
        name: "Excel.TimelineStyle.name",
        description: "Specifies the name of the timeline style.",
        kind: "Property",
        signature: "Excel.TimelineStyle.name: string",
        examples: [],
      },
      {
        name: "Excel.TimelineStyle.readOnly",
        description: "Specifies if this `TimelineStyle` object is read-only.",
        kind: "Property",
        signature: "Excel.TimelineStyle.readOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.TimelineStyle.delete",
        description: "Deletes the table style.",
        kind: "Method",
        signature: "Excel.TimelineStyle.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.TimelineStyle.duplicate",
        description:
          "Creates a duplicate of this timeline style with copies of all the style elements.",
        kind: "Method",
        signature: "Excel.TimelineStyle.duplicate => () => Excel.TimelineStyle",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TimelineStyleCollection",
    apiList: [
      {
        name: "Excel.TimelineStyleCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.TimelineStyleCollection.items: TimelineStyle[]",
        examples: [],
      },
      {
        name: "Excel.TimelineStyleCollection.add",
        description: "Creates a blank `TimelineStyle` with the specified name.",
        kind: "Method",
        signature:
          "Excel.TimelineStyleCollection.add => (name: string, makeUniqueName?: boolean) => Excel.TimelineStyle",
        examples: [],
      },
      {
        name: "Excel.TimelineStyleCollection.getCount",
        description: "Gets the number of timeline styles in the collection.",
        kind: "Method",
        signature:
          "Excel.TimelineStyleCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.TimelineStyleCollection.getDefault",
        description: "Gets the default timeline style for the parent object's scope.",
        kind: "Method",
        signature: "Excel.TimelineStyleCollection.getDefault => () => Excel.TimelineStyle",
        examples: [],
      },
      {
        name: "Excel.TimelineStyleCollection.getItem",
        description: "Gets a `TimelineStyle` by name.",
        kind: "Method",
        signature: "Excel.TimelineStyleCollection.getItem => (name: string) => Excel.TimelineStyle",
        examples: [],
      },
      {
        name: "Excel.TimelineStyleCollection.setDefault",
        description: "Sets the default timeline style for use in the parent object's scope.",
        kind: "Method",
        signature:
          "Excel.TimelineStyleCollection.setDefault => (newDefaultStyle: TimelineStyle | string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.TopBottomConditionalFormat",
    apiList: [
      {
        name: "Excel.TopBottomConditionalFormat.format",
        description:
          "Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.",
        kind: "Property",
        signature: "Excel.TopBottomConditionalFormat.format: Excel.ConditionalRangeFormat",
        examples: ['conditionalFormat.topBottom.format.fill.color = "green";'],
      },
      {
        name: "Excel.TopBottomConditionalFormat.rule",
        description: "The criteria of the top/bottom conditional format.",
        kind: "Property",
        signature: "Excel.TopBottomConditionalFormat.rule: Excel.ConditionalTopBottomRule",
        examples: ['conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems" };'],
      },
    ],
  },
  {
    objName: "Excel.UserActivity",
    apiList: [
      {
        name: "Excel.UserActivity.activityId",
        description:
          "The ID for the user activity. This has a 1:1 relationship with the revision ID in Excel client.",
        kind: "Property",
        signature: "Excel.UserActivity.activityId: number",
        examples: [],
      },
      {
        name: "Excel.UserActivity.activityType",
        description: "Type of activity.",
        kind: "Property",
        signature:
          'Excel.UserActivity.activityType: UserActivityType | "None" | "InsertSheet" | "DeleteSheet" | "RenameSheet" | "ChangeCell" | "InsertRow" | "InsertColumn" | "DeleteRow" | "DeleteColumn" | ... 11 more ... | "GenericEdit"',
        examples: [],
      },
      {
        name: "Excel.UserActivity.author",
        description: "Author who created the activity.",
        kind: "Property",
        signature: "Excel.UserActivity.author: Identity",
        examples: [],
      },
      {
        name: "Excel.UserActivity.authorEmail",
        description: "Email address of the author who created the activity.",
        kind: "Property",
        signature: "Excel.UserActivity.authorEmail: string",
        examples: [],
      },
      {
        name: "Excel.UserActivity.createdDateTime",
        description: "The time when the activity was created.",
        kind: "Property",
        signature: "Excel.UserActivity.createdDateTime: Date",
        examples: [],
      },
      {
        name: "Excel.UserActivity.guid",
        description: "Unique identifier of the activity.",
        kind: "Property",
        signature: "Excel.UserActivity.guid: string",
        examples: [],
      },
      {
        name: "Excel.UserActivity.highlightRangeAreas",
        description: "The range affected by the activity. Can be a discontiguous range.",
        kind: "Property",
        signature: "Excel.UserActivity.highlightRangeAreas: RangeAreas",
        examples: [],
      },
      {
        name: "Excel.UserActivity.locationDeleted",
        description: "Boolean to indicate deleted location activity card type.",
        kind: "Property",
        signature: "Excel.UserActivity.locationDeleted: boolean",
        examples: [],
      },
      {
        name: "Excel.UserActivity.rangeAddress",
        description:
          "Represents the address of the range where the activity happened. This is a contiguous range that contains all the ranges affected by the activity.",
        kind: "Property",
        signature: "Excel.UserActivity.rangeAddress: string",
        examples: [],
      },
      {
        name: "Excel.UserActivity.sheetName",
        description: "The sheet name where the activity happened.",
        kind: "Property",
        signature: "Excel.UserActivity.sheetName: string",
        examples: [],
      },
      {
        name: "Excel.UserActivity.valueChangeData",
        description: "The list of cell value changes associated with the activity.",
        kind: "Property",
        signature: "Excel.UserActivity.valueChangeData: UserActivityCellValueChangeData",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.UserActivityCellValueChangeData",
    apiList: [
      {
        name: "Excel.UserActivityCellValueChangeData.allAvailable",
        description: "Flag denoting if all the value changes are available.",
        kind: "Property",
        signature: "Excel.UserActivityCellValueChangeData.allAvailable: boolean",
        examples: [],
      },
      {
        name: "Excel.UserActivityCellValueChangeData.valueChanges",
        description:
          "UserActivityCellValueChange the contains list of cell value changes associated with the activity.",
        kind: "Property",
        signature:
          "Excel.UserActivityCellValueChangeData.valueChanges: UserActivityCellValueChange[]",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.UserActivityCollection",
    apiList: [
      {
        name: "Excel.UserActivityCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.UserActivityCollection.items: UserActivity[]",
        examples: [],
      },
      {
        name: "Excel.UserActivityCollection.getCount",
        description: "Gets the number of activities in the collection.",
        kind: "Method",
        signature:
          "Excel.UserActivityCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.UserActivityCollection.getItemAt",
        description: "Gets the `UserActivity` object by its index in the collection.",
        kind: "Method",
        signature:
          "Excel.UserActivityCollection.getItemAt => (index: number) => Excel.UserActivity",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.UserActivityFilter",
    apiList: [
      {
        name: "Excel.UserActivityFilter.rangeAddress",
        description:
          "A range address. This filters the activities to only activities from this range.",
        kind: "Property",
        signature: "Excel.UserActivityFilter.rangeAddress: string",
        examples: [],
      },
      {
        name: "Excel.UserActivityFilter.sheetName",
        description:
          "A worksheet name. This filters the activities to only activities from this worksheet.",
        kind: "Property",
        signature: "Excel.UserActivityFilter.sheetName: string",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ValueErrorCellValue",
    apiList: [
      {
        name: "Excel.ValueErrorCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.ValueErrorCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.ValueErrorCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.ValueErrorCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.ValueErrorCellValue.errorSubType",
        description: "Represents the type of `ValueErrorCellValue`.",
        kind: "Property",
        signature:
          'Excel.ValueErrorCellValue.errorSubType: "Unknown" | ValueErrorCellValueSubType | "VlookupColumnIndexLessThanOne" | "VlookupResultNotFound" | "HlookupRowIndexLessThanOne" | "HlookupResultNotFound" | ... 14 more ... | "LambdaWrongParamCount"',
        examples: [],
      },
      {
        name: "Excel.ValueErrorCellValue.errorType",
        description: "Represents the type of `ErrorCellValue`.",
        kind: "Property",
        signature: 'Excel.ValueErrorCellValue.errorType: ErrorCellValueType.value | "Value"',
        examples: [],
      },
      {
        name: "Excel.ValueErrorCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.ValueErrorCellValue.type: CellValueType.error | "Error"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.ValueTypeNotAvailableCellValue",
    apiList: [
      {
        name: "Excel.ValueTypeNotAvailableCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature:
          'Excel.ValueTypeNotAvailableCellValue.basicType: RangeValueType | "Error" | "Boolean" | "Double" | "Empty" | "String"',
        examples: [],
      },
      {
        name: "Excel.ValueTypeNotAvailableCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value.",
        kind: "Property",
        signature: "Excel.ValueTypeNotAvailableCellValue.basicValue: string | number | boolean",
        examples: [],
      },
      {
        name: "Excel.ValueTypeNotAvailableCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature:
          'Excel.ValueTypeNotAvailableCellValue.type: CellValueType.notAvailable | "NotAvailable"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Visual",
    apiList: [
      {
        name: "Excel.Visual.id",
        description: "The unique ID of this visual instance.",
        kind: "Property",
        signature: "Excel.Visual.id: string",
        examples: [],
      },
      {
        name: "Excel.Visual.isSupportedInVisualTaskpane",
        description:
          "Represents if the visual is supported in the new Excel on the web chart format task pane.",
        kind: "Property",
        signature: "Excel.Visual.isSupportedInVisualTaskpane: boolean",
        examples: [],
      },
      {
        name: "Excel.Visual.properties",
        description: "Gets all properties of the visual.",
        kind: "Property",
        signature: "Excel.Visual.properties: VisualPropertyCollection",
        examples: [],
      },
      {
        name: "Excel.Visual.addChildProperty",
        description:
          "Adds a new property to a parent collection. Only valid for properties of the type `VisualPropertyType.Collection`.",
        kind: "Method",
        signature:
          "Excel.Visual.addChildProperty => (parentCollectionName: string, attributes?: any) => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.Visual.changeDataSource",
        description: "Change the data source of the visual.",
        kind: "Method",
        signature:
          "Excel.Visual.changeDataSource => (dataSourceType: string, dataSourceContent: string) => void",
        examples: [],
      },
      {
        name: "Excel.Visual.delete",
        description: "Deletes the visual.",
        kind: "Method",
        signature: "Excel.Visual.delete => () => void",
        examples: [],
      },
      {
        name: "Excel.Visual.deserializeProperties",
        description: "Recursively modify UO properties.",
        kind: "Method",
        signature: "Excel.Visual.deserializeProperties => (json: string) => void",
        examples: [],
      },
      {
        name: "Excel.Visual.getChildProperties",
        description: "Gets the child properties of the specific parent property ID.",
        kind: "Method",
        signature:
          "Excel.Visual.getChildProperties => (parentPropId?: string, levelsToTraverse?: number) => Excel.VisualPropertyCollection",
        examples: [],
      },
      {
        name: "Excel.Visual.getDataConfig",
        description:
          "Gets the visual's data configuration. Data configuration JSON is generated in ChartVisual.cpp GetDataConfigJson method.",
        kind: "Method",
        signature: "Excel.Visual.getDataConfig => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.Visual.getDataControllerClient",
        description: "Gets the data controller client for the visual.",
        kind: "Method",
        signature: "Excel.Visual.getDataControllerClient => () => Excel.DataControllerClient",
        examples: [],
      },
      {
        name: "Excel.Visual.getDataFieldAssignments",
        description:
          "Data field assignments are named sets of fields, such as category fields or value fields. In a data field assignment of category fields, the fields contain the ranges for each category entry.",
        kind: "Method",
        signature:
          "Excel.Visual.getDataFieldAssignments => () => OfficeExtension.ClientResult<Excel.VisualDataFieldAssignment[]>",
        examples: [],
      },
      {
        name: "Excel.Visual.getDataSource",
        description:
          'Gets a string representing the visual\'s current data source (e.g., "Sheet1!$C$5:$D$7").',
        kind: "Method",
        signature: "Excel.Visual.getDataSource => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.Visual.getElementChildProperties",
        description: "Gets the child properties of the specific parent linked to cookie.",
        kind: "Method",
        signature:
          "Excel.Visual.getElementChildProperties => (elementId: number, index: number, levelsToTraverse?: number) => Excel.VisualPropertyCollection",
        examples: [],
      },
      {
        name: "Excel.Visual.getProperty",
        description: "GetProperty",
        kind: "Method",
        signature:
          "Excel.Visual.getProperty => (propName: string) => OfficeExtension.ClientResult<any>",
        examples: [],
      },
      {
        name: "Excel.Visual.modifyDataConfig",
        description:
          "Modifies the visual's data configuration. Data modification JSON is consumed in ChartVisual.cpp ApplyDataConfigJson method.",
        kind: "Method",
        signature: "Excel.Visual.modifyDataConfig => (configModification: string) => void",
        examples: [],
      },
      {
        name: "Excel.Visual.removeChildProperty",
        description:
          "Removes a property from the parent collection. Only valid for properties of type `VisualPropertyType.Collection`.",
        kind: "Method",
        signature:
          "Excel.Visual.removeChildProperty => (parentCollectionName: string, index: number) => void",
        examples: [],
      },
      {
        name: "Excel.Visual.serializeProperties",
        description: "Recursively serialize UO properties.",
        kind: "Method",
        signature: "Excel.Visual.serializeProperties => () => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.Visual.setProperty",
        description: "SetProperty",
        kind: "Method",
        signature: "Excel.Visual.setProperty => (propName: string, value: any) => void",
        examples: [],
      },
      {
        name: "Excel.Visual.setPropertyToDefault",
        description: "Returns `true` when the property's value is currently the default.",
        kind: "Method",
        signature: "Excel.Visual.setPropertyToDefault => (propName: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.VisualCollection",
    apiList: [
      {
        name: "Excel.VisualCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.VisualCollection.items: Visual[]",
        examples: [],
      },
      {
        name: "Excel.VisualCollection.add",
        description: "Creates a new visual.",
        kind: "Method",
        signature:
          "Excel.VisualCollection.add => (visualDefinitionGuid: string, dataSourceType?: string, dataSourceContent?: string) => Excel.Visual",
        examples: [],
      },
      {
        name: "Excel.VisualCollection.bootstrapAgaveVisual",
        description:
          "Creates a new agave visual from the calling content add-in. Similar to initializing an agave visual with `Add()`, except the add-in instance already exists. Additionally, registers the `AgaveVisualUpdate` event.",
        kind: "Method",
        signature: "Excel.VisualCollection.bootstrapAgaveVisual => () => void",
        examples: [],
      },
      {
        name: "Excel.VisualCollection.getCount",
        description: "Returns the number of visuals in the worksheet.",
        kind: "Method",
        signature: "Excel.VisualCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.VisualCollection.getDefinitions",
        description: "Gets all visual definitions.",
        kind: "Method",
        signature:
          "Excel.VisualCollection.getDefinitions => () => OfficeExtension.ClientResult<Excel.VisualDefinition[]>",
        examples: [],
      },
      {
        name: "Excel.VisualCollection.getPreview",
        description: "Get the preview of a visual.",
        kind: "Method",
        signature:
          "Excel.VisualCollection.getPreview => (visualDefinitionGuid: string, width: number, height: number, dpi: number) => OfficeExtension.ClientResult<string>",
        examples: [],
      },
      {
        name: "Excel.VisualCollection.getSelectedOrNullObject",
        description:
          "Gets the selected visual, if and only if one visual is selected. If no visual is selected, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.VisualCollection.getSelectedOrNullObject => () => Excel.Visual",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.VisualProperty",
    apiList: [
      {
        name: "Excel.VisualProperty.expandableUI",
        description: "Returns `true` if the property should be expandable in the UI.",
        kind: "Property",
        signature: "Excel.VisualProperty.expandableUI: boolean",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.hasDefault",
        description: "Returns true when a default value for this property exists",
        kind: "Property",
        signature: "Excel.VisualProperty.hasDefault: boolean",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.hideMeButShowChildrenUI",
        description:
          "Returns `true` if the property should be hidden in the UI. Its children will still be shown in the UI.",
        kind: "Property",
        signature: "Excel.VisualProperty.hideMeButShowChildrenUI: boolean",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.id",
        description: "Returns the property ID.",
        kind: "Property",
        signature: "Excel.VisualProperty.id: string",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.index",
        description:
          "The zero-index value at which the property is present in the parent collection. Only valid for properties that are children of `VisualPropertyType.Collection`. Returns `null` otherwise.",
        kind: "Property",
        signature: "Excel.VisualProperty.index: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.isDefault",
        description: "Returns true when the property's value is currently the default",
        kind: "Property",
        signature: "Excel.VisualProperty.isDefault: boolean",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.localizedName",
        description: "Returns the property localized name.",
        kind: "Property",
        signature: "Excel.VisualProperty.localizedName: string",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.localizedOptions",
        description:
          "Returns the localized property options for `IEnumProperty` only. If property type isn't an enum, it returns `null`.",
        kind: "Property",
        signature: "Excel.VisualProperty.localizedOptions: string[]",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.max",
        description:
          "Returns the maximum value of the property. Only valid for `INumericProperty` properties. Returns `null` if it's invalid.",
        kind: "Property",
        signature: "Excel.VisualProperty.max: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.maxSize",
        description:
          "Maximum size of the property. Only valid for `VisualPropertyType.Collection`. Returns `null` if it's invalid.",
        kind: "Property",
        signature: "Excel.VisualProperty.maxSize: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.min",
        description:
          "Returns the minimum value of the property. Only valid for `INumericProperty` properties. Returns `null` if it's invalid.",
        kind: "Property",
        signature: "Excel.VisualProperty.min: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.minSize",
        description:
          "Minimum size of the property. Only valid for `VisualPropertyType.Collection`. Returns `null` if it's invalid.",
        kind: "Property",
        signature: "Excel.VisualProperty.minSize: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.nextPropOnSameLine",
        description: "Returns `true` if the next property should be on the same line in the UI.",
        kind: "Property",
        signature: "Excel.VisualProperty.nextPropOnSameLine: boolean",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.options",
        description:
          "Returns the property options for `IEnumProperty` only. If the property type isn't an enum, it returns `null`.",
        kind: "Property",
        signature: "Excel.VisualProperty.options: string[]",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.parentName",
        description:
          "Name of the parent property. Only valid for properties that are children of `VisualPropertyType.Collection`. Returns `null` otherwise.",
        kind: "Property",
        signature: "Excel.VisualProperty.parentName: string",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.showResetUI",
        description: "Returns `true` if a reset button for the property should be shown in the UI.",
        kind: "Property",
        signature: "Excel.VisualProperty.showResetUI: boolean",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.size",
        description:
          "Size of the property. Only valid for `VisualPropertyType.Collection`. Returns `null` if it's invalid.",
        kind: "Property",
        signature: "Excel.VisualProperty.size: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.stepSize",
        description:
          "Returns the step size of the property. Only valid for `INumericProperty` properties. Returns `null` if it's invalid.",
        kind: "Property",
        signature: "Excel.VisualProperty.stepSize: number",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.type",
        description: "Returns the property type.",
        kind: "Property",
        signature:
          'Excel.VisualProperty.type: "Double" | "String" | VisualPropertyType | "Object" | "Collection" | "Int" | "Bool" | "Enum" | "Color"',
        examples: [],
      },
      {
        name: "Excel.VisualProperty.value",
        description: "Returns the property value.",
        kind: "Property",
        signature: "Excel.VisualProperty.value: any",
        examples: [],
      },
      {
        name: "Excel.VisualProperty.getBoolMetaProperty",
        description:
          "Returns `true` if the visual property's boolean meta-property is set. The type of meta property",
        kind: "Method",
        signature:
          'Excel.VisualProperty.getBoolMetaProperty => { (metaProp: BoolMetaPropertyType): OfficeExtension.ClientResult<boolean>; (metaProp: "WriteOnly" | "ReadOnly" | "HideEntireSubtreeUI" | ... 8 more ... | "Untransferable"): OfficeExtension.ClientResult<...>; (metaProp: string): OfficeExtension.ClientResult<boolean>; }',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.VisualPropertyCollection",
    apiList: [
      {
        name: "Excel.VisualPropertyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.VisualPropertyCollection.items: VisualProperty[]",
        examples: [],
      },
      {
        name: "Excel.VisualPropertyCollection.getCount",
        description: "Returns the number of properties in the collection.",
        kind: "Method",
        signature:
          "Excel.VisualPropertyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.VisualPropertyCollection.getItem",
        description: "Returns a property at the given index.",
        kind: "Method",
        signature:
          "Excel.VisualPropertyCollection.getItem => (index: number) => Excel.VisualProperty",
        examples: [],
      },
      {
        name: "Excel.VisualPropertyCollection.getItemAt",
        description: "Returns a property at the given index.",
        kind: "Method",
        signature:
          "Excel.VisualPropertyCollection.getItemAt => (index: number) => Excel.VisualProperty",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.VisualTracker",
    apiList: [
      {
        name: "Excel.VisualTracker.id",
        description: "ID for the visual tracker.",
        kind: "Property",
        signature: "Excel.VisualTracker.id: string",
        examples: [],
      },
      {
        name: "Excel.VisualTracker.requestTrackingAlteration",
        description:
          "Make a request to change the tracking being done on visuals. The JSON encoding the request is produced by the biplat.uniformobjects library, as the return value from `VisualsMirror.createTrackingAlterationRequestJson()`.",
        kind: "Method",
        signature:
          "Excel.VisualTracker.requestTrackingAlteration => (requestSourceName: string, trackingAlterationRequestJson: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WebImageCellValue",
    apiList: [
      {
        name: "Excel.WebImageCellValue.address",
        description:
          "Represents the URL from which the image will be downloaded. This image must be hosted on a server that supports HTTPS.",
        kind: "Property",
        signature: "Excel.WebImageCellValue.address: string",
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.altText",
        description:
          "Represents the alternate text that can be used in accessibility scenarios to describe what the image represents.",
        kind: "Property",
        signature: "Excel.WebImageCellValue.altText: string",
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.attribution",
        description:
          "Represents attribution information to describe the source and license requirements for using this image.",
        kind: "Property",
        signature: "Excel.WebImageCellValue.attribution: CellValueAttributionAttributes[]",
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.basicType",
        description:
          "Represents the value that would be returned by `Range.valueTypes` for a cell with this value.",
        kind: "Property",
        signature: 'Excel.WebImageCellValue.basicType: RangeValueType.error | "Error"',
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.basicValue",
        description:
          "Represents the value that would be returned by `Range.values` for a cell with this value. When accessed through a `valuesAsJson` property, this string value aligns with the en-US locale. When accessed through a `valuesAsJsonLocal` property, this string value aligns with the user's display locale.",
        kind: "Property",
        signature: "Excel.WebImageCellValue.basicValue: string",
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.provider",
        description:
          "Represents information that describes the entity or individual who provided the image. This information can be used for branding in image cards.",
        kind: "Property",
        signature: "Excel.WebImageCellValue.provider: CellValueProviderAttributes",
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.relatedImagesAddress",
        description:
          "Represents the URL of a webpage with images that are considered related to this `WebImageCellValue`.",
        kind: "Property",
        signature: "Excel.WebImageCellValue.relatedImagesAddress: string",
        examples: [],
      },
      {
        name: "Excel.WebImageCellValue.type",
        description: "Represents the type of this cell value.",
        kind: "Property",
        signature: 'Excel.WebImageCellValue.type: CellValueType.webImage | "WebImage"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Workbook",
    apiList: [
      {
        name: "Excel.Workbook.application",
        description: "Represents the Excel application instance that contains this workbook.",
        kind: "Property",
        signature: "Excel.Workbook.application: Excel.Application",
        examples: [
          "let app = workbook.application;",
          "workbook.application.calculate(Excel.CalculationType.full);",
          'workbook.application.calculate("Full");',
          "const localDecimalSeparator = workbook.application.decimalSeparator;",
          "const localThousandsSeparator = workbook.application.thousandsSeparator;",
          "const systemDecimalSeparator = workbook.application.cultureInfo.numberFormat.numberDecimalSeparator;",
          "const systemThousandsSeparator = workbook.application.cultureInfo.numberFormat.numberGroupSeparator;",
          "const application = workbook.application;",
          "workbook.application.calculationMode = Excel.CalculationMode.manual;",
          '"Current calculation mode: " + workbook.application.calculationMode;',
          "workbook.application.calculate(Excel.CalculationType.recalculate);",
          "const systemLongDatePattern = workbook.application.cultureInfo.datetimeFormat.longDatePattern;",
          "const systemShortDatePattern = workbook.application.cultureInfo.datetimeFormat.shortDatePattern;",
          "const systemDateSeparator = workbook.application.cultureInfo.datetimeFormat.dateSeparator;",
          "const systemLongTimePattern = workbook.application.cultureInfo.datetimeFormat.longTimePattern;",
          "const systemTimeSeparator = workbook.application.cultureInfo.datetimeFormat.timeSeparator;",
        ],
      },
      {
        name: "Excel.Workbook.autoSave",
        description: "Specifies if the workbook is in AutoSave mode.",
        kind: "Property",
        signature: "Excel.Workbook.autoSave: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.bindings",
        description: "Represents a collection of bindings that are part of the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.bindings: Excel.BindingCollection",
        examples: [
          "const binding = workbook.bindings.getItemAt(0);",
          "const lastPosition = workbook.bindings.count - 1;",
          "const binding = workbook.bindings.getItemAt(lastPosition);",
        ],
      },
      {
        name: "Excel.Workbook.calculationEngineVersion",
        description: "Returns a number about the version of Excel Calculation Engine.",
        kind: "Property",
        signature: "Excel.Workbook.calculationEngineVersion: number",
        examples: [],
      },
      {
        name: "Excel.Workbook.chartDataPointTrack",
        description:
          "True if all charts in the workbook are tracking the actual data points to which they are attached. False if the charts track the index of the data points.",
        kind: "Property",
        signature: "Excel.Workbook.chartDataPointTrack: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.comments",
        description: "Represents a collection of comments associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.comments: Excel.CommentCollection",
        examples: [
          "let comments = workbook.comments;",
          "let comment = workbook.comments.getItemAt(0);",
          'workbook.comments.getItemByCell("MyWorksheet!A2").delete();',
          "workbook.comments.getItemAt(0).resolved = true;",
          'let comment = workbook.comments.getItemByCell("MyWorksheet!A2");',
          'workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);',
          'workbook.comments.getItemByCell("Comments!A2").delete();',
          'const comment = workbook.comments.getItemByCell("Comments!A2");',
        ],
      },
      {
        name: "Excel.Workbook.formulaReferenceStyle",
        description:
          "Represents the formula reference style used by the workbook. R1C1 formula reference style is only available in Excel on Windows and Mac. It's not available in Excel on the web.",
        kind: "Property",
        signature: 'Excel.Workbook.formulaReferenceStyle: FormulaReferenceStyle | "A1" | "R1C1"',
        examples: [],
      },
      {
        name: "Excel.Workbook.functions",
        description:
          "Represents a collection of worksheet functions that can be used for computation.",
        kind: "Property",
        signature: "Excel.Workbook.functions: Excel.Functions",
        examples: ['let unitSoldInNov = workbook.functions.vlookup("Wrench", range, 2, false);'],
      },
      {
        name: "Excel.Workbook.guidedReapply",
        description: "Returns the `GuidedReapplyManager` object associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.guidedReapply: GuidedReapplyManager",
        examples: [],
      },
      {
        name: "Excel.Workbook.isDirty",
        description:
          "Specifies if changes have been made since the workbook was last saved. You can set this property to `true` if you want to close a modified workbook without either saving it or being prompted to save it.",
        kind: "Property",
        signature: "Excel.Workbook.isDirty: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.lineageActivities",
        description: "Returns the lineageActivityCollection object associated with workbook.",
        kind: "Property",
        signature: "Excel.Workbook.lineageActivities: LineageActivityCollection",
        examples: [],
      },
      {
        name: "Excel.Workbook.name",
        description: "Gets the workbook name.",
        kind: "Property",
        signature: "Excel.Workbook.name: string",
        examples: [],
      },
      {
        name: "Excel.Workbook.names",
        description:
          "Represents a collection of workbook-scoped named items (named ranges and constants).",
        kind: "Property",
        signature: "Excel.Workbook.names: Excel.NamedItemCollection",
        examples: [
          "const names = workbook.names;",
          "const nameditem = workbook.names.getItem(sheetName);",
        ],
      },
      {
        name: "Excel.Workbook.pivotTables",
        description: "Represents a collection of PivotTables associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.pivotTables: Excel.PivotTableCollection",
        examples: [
          'workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");',
        ],
      },
      {
        name: "Excel.Workbook.pivotTableStyles",
        description: "Represents a collection of PivotTableStyles associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.pivotTableStyles: PivotTableStyleCollection",
        examples: [],
      },
      {
        name: "Excel.Workbook.previouslySaved",
        description: "Specifies if the workbook has ever been saved locally or online.",
        kind: "Property",
        signature: "Excel.Workbook.previouslySaved: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.properties",
        description: "Gets the workbook properties.",
        kind: "Property",
        signature: "Excel.Workbook.properties: Excel.DocumentProperties",
        examples: ["let docProperties = workbook.properties;"],
      },
      {
        name: "Excel.Workbook.protection",
        description: "Returns the protection object for a workbook.",
        kind: "Property",
        signature: "Excel.Workbook.protection: Excel.WorkbookProtection",
        examples: ["workbook.protection.protect();"],
      },
      {
        name: "Excel.Workbook.readOnly",
        description: "Returns `true` if the workbook is open in read-only mode.",
        kind: "Property",
        signature: "Excel.Workbook.readOnly: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.settings",
        description: "Represents a collection of settings associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.settings: Excel.SettingCollection",
        examples: ["let settings = workbook.settings;"],
      },
      {
        name: "Excel.Workbook.showPivotFieldList",
        description:
          "Specifies whether the PivotTable's field list pane is shown at the workbook level.",
        kind: "Property",
        signature: "Excel.Workbook.showPivotFieldList: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.slicers",
        description: "Represents a collection of slicers associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.slicers: Excel.SlicerCollection",
        examples: [
          'let slicer = workbook.slicers.getItem("Fruit Slicer");',
          'const slicer = workbook.slicers.getItem("Fruit Slicer");',
        ],
      },
      {
        name: "Excel.Workbook.slicerStyles",
        description: "Represents a collection of SlicerStyles associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.slicerStyles: SlicerStyleCollection",
        examples: [],
      },
      {
        name: "Excel.Workbook.styles",
        description: "Represents a collection of styles associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.styles: Excel.StyleCollection",
        examples: [
          'let style = workbook.styles.getItem("Diagonal Orientation Style");',
          'let style = workbook.styles.getItem("Normal");',
          "let styles = workbook.styles;",
        ],
      },
      {
        name: "Excel.Workbook.tables",
        description: "Represents a collection of tables associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.tables: Excel.TableCollection",
        examples: [
          'const table = workbook.tables.add("Sheet1!A1:E7", true);',
          "const tables = workbook.tables;",
        ],
      },
      {
        name: "Excel.Workbook.tableStyles",
        description: "Represents a collection of TableStyles associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.tableStyles: TableStyleCollection",
        examples: [],
      },
      {
        name: "Excel.Workbook.tasks",
        description: "Returns a collection of tasks that are present in the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.tasks: DocumentTaskCollection",
        examples: [],
      },
      {
        name: "Excel.Workbook.timelineStyles",
        description: "Represents a collection of TimelineStyles associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.timelineStyles: TimelineStyleCollection",
        examples: [],
      },
      {
        name: "Excel.Workbook.use1904DateSystem",
        description: "True if the workbook uses the 1904 date system.",
        kind: "Property",
        signature: "Excel.Workbook.use1904DateSystem: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.usePrecisionAsDisplayed",
        description:
          "True if calculations in this workbook will be done using only the precision of the numbers as they're displayed. Data will permanently lose accuracy when switching this property from `false` to `true`.",
        kind: "Property",
        signature: "Excel.Workbook.usePrecisionAsDisplayed: boolean",
        examples: [],
      },
      {
        name: "Excel.Workbook.worksheets",
        description: "Represents a collection of worksheets associated with the workbook.",
        kind: "Property",
        signature: "Excel.Workbook.worksheets: Excel.WorksheetCollection",
        examples: [
          "const activeWorksheet = workbook.worksheets.getActiveWorksheet();",
          'let rangeToAnalyze = workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");',
          'let rangeToPlacePivot = workbook.worksheets.getItem("PivotWorksheet").getRange("A2");',
          'workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
          "let firstSheet = workbook.worksheets.getFirst();",
          "let lastSheet = workbook.worksheets.getLast();",
          "let sheets = workbook.worksheets;",
          'const sheet = workbook.worksheets.getItemOrNullObject("Sample");',
          "const range = workbook.worksheets.getItem(sheetName).getRange(rangeSelection);",
          'const chart = workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");',
          'const lastPosition = workbook.worksheets.getItem("Sheet1").charts.count - 1;',
          'const chart = workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);',
          'workbook.worksheets.getItemOrNullObject("Sample").delete();',
          'const sheet = workbook.worksheets.add("Sample");',
          'const nameSourceRange = workbook.worksheets.getItem("Names").getRange("A1:A3");',
          'const rangeToAnalyze = workbook.worksheets.getItem("Data").getRange("A1:E21");',
          'const rangeToPlacePivot = workbook.worksheets.getItem("Pivot").getRange("A2");',
          'workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
          'workbook.worksheets.getItemOrNullObject("Shapes").delete();',
          'const sheet = workbook.worksheets.add("Shapes");',
          "const sheets = workbook.worksheets;",
          "const worksheet = workbook.worksheets.add(wSheetName);",
        ],
      },
      {
        name: "Excel.Workbook.close",
        description: "Close current workbook.",
        kind: "Method",
        signature: "Excel.Workbook.close(closeBehavior?: Excel.CloseBehavior): void",
        examples: [
          "workbook.close(Excel.CloseBehavior.skipSave);",
          "workbook.close(Excel.CloseBehavior.save);",
        ],
      },
      {
        name: "Excel.Workbook.focus",
        description:
          "Sets focus on the workbook. This will cause the grid or the currently active object to receive keyboard events.",
        kind: "Method",
        signature: "Excel.Workbook.focus => () => void",
        examples: [],
      },
      {
        name: "Excel.Workbook.getActiveCell",
        description: "Gets the currently active cell from the workbook.",
        kind: "Method",
        signature: "Excel.Workbook.getActiveCell() => Excel.Range",
        examples: [
          "let activeCell = workbook.getActiveCell();",
          "const cell = workbook.getActiveCell();",
          "const activeCell = workbook.getActiveCell();",
          "let activeCell = myWorkbook.getActiveCell();",
        ],
      },
      {
        name: "Excel.Workbook.getActiveChart",
        description:
          "Gets the currently active chart in the workbook. If there is no active chart, an `ItemNotFound` exception is thrown.",
        kind: "Method",
        signature: "Excel.Workbook.getActiveChart() => Excel.Chart",
        examples: [],
      },
      {
        name: "Excel.Workbook.getActiveChartOrNullObject",
        description:
          "Gets the currently active chart in the workbook. If there is no active chart, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Workbook.getActiveChartOrNullObject => () => Excel.Chart",
        examples: [],
      },
      {
        name: "Excel.Workbook.getActiveSlicer",
        description:
          "Gets the currently active slicer in the workbook. If there is no active slicer, an `ItemNotFound` exception is thrown.",
        kind: "Method",
        signature: "Excel.Workbook.getActiveSlicer => () => Excel.Slicer",
        examples: [],
      },
      {
        name: "Excel.Workbook.getActiveSlicerOrNullObject",
        description:
          "Gets the currently active slicer in the workbook. If there is no active slicer, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.Workbook.getActiveSlicerOrNullObject => () => Excel.Slicer",
        examples: [],
      },
      {
        name: "Excel.Workbook.getIsActiveCollabSession",
        description:
          "Returns `true` if the workbook is being edited by multiple users (through co-authoring). Please be aware there might be some delay between when the workbook status changes and when the changes are reflected on the result of the method.",
        kind: "Method",
        signature:
          "Excel.Workbook.getIsActiveCollabSession => () => OfficeExtension.ClientResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.Workbook.getLinkedEntityCellValue",
        description: "Returns a `LinkedEntityCellValue` based on the provided `LinkedEntityId`.",
        kind: "Method",
        signature:
          "Excel.Workbook.getLinkedEntityCellValue => (linkedEntityCellValueId: LinkedEntityId) => OfficeExtension.ClientResult<LinkedEntityCellValue>",
        examples: [],
      },
      {
        name: "Excel.Workbook.getSelectedRange",
        description:
          "Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.",
        kind: "Method",
        signature: "Excel.Workbook.getSelectedRange() => Excel.Range",
        examples: ["const selectedRange = workbook.getSelectedRange();"],
      },
      {
        name: "Excel.Workbook.getSelectedRanges",
        description:
          "Gets the currently selected one or more ranges from the workbook. Unlike `getSelectedRange()`, this method returns a `RangeAreas` object that represents all the selected ranges.",
        kind: "Method",
        signature: "Excel.Workbook.getSelectedRanges() => Excel.RangeAreas",
        examples: ["const selectedRanges = workbook.getSelectedRanges();"],
      },
      {
        name: "Excel.Workbook.getThemeColors",
        description:
          "Provides a list of theme colors in Excel, based on the theme/color scheme applied to the document. These theme colors will be used to populate the theme colors palette in the color picker menu.",
        kind: "Method",
        signature: "Excel.Workbook.getThemeColors => () => OfficeExtension.ClientResult<number[]>",
        examples: [],
      },
      {
        name: "Excel.Workbook.save",
        description: "Save current workbook.",
        kind: "Method",
        signature: "Excel.Workbook.save(saveBehavior?: Excel.SaveBehavior): void",
        examples: [
          "workbook.save(Excel.SaveBehavior.prompt);",
          "workbook.save(Excel.SaveBehavior.save);",
        ],
      },
    ],
  },
  {
    objName: "Excel.WorkbookCreated",
    apiList: [
      {
        name: "Excel.WorkbookCreated.open",
        description: "Open the workbook.",
        kind: "Method",
        signature: "Excel.WorkbookCreated.open => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorkbookProtection",
    apiList: [
      {
        name: "Excel.WorkbookProtection.protected",
        description: "Specifies if the workbook is protected.",
        kind: "Property",
        signature: "Excel.WorkbookProtection.protected: boolean",
        examples: [
          "if (!workbook.protection.protected) {\n    workbook.protection.protect();\n  }",
        ],
      },
      {
        name: "Excel.WorkbookProtection.protect",
        description: "Protects a workbook. Fails if the workbook has been protected.",
        kind: "Method",
        signature: "Excel.WorkbookProtection.protect(password?: string) => void",
        examples: ["workbook.protection.protect();"],
      },
      {
        name: "Excel.WorkbookProtection.unprotect",
        description: "Unprotects a workbook.",
        kind: "Method",
        signature: "Excel.WorkbookProtection.unprotect => (password?: string) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorkbookRangeAreas",
    apiList: [
      {
        name: "Excel.WorkbookRangeAreas.addresses",
        description:
          'Returns an array of addresses in A1-style. Address values contain the worksheet name for each rectangular block of cells (e.g., "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.',
        kind: "Property",
        signature: "Excel.WorkbookRangeAreas.addresses: string[]",
        examples: [],
      },
      {
        name: "Excel.WorkbookRangeAreas.areas",
        description:
          "Returns the `RangeAreasCollection` object. Each `RangeAreas` in the collection represent one or more rectangle ranges in one worksheet.",
        kind: "Property",
        signature: "Excel.WorkbookRangeAreas.areas: Excel.RangeAreasCollection",
        examples: [],
      },
      {
        name: "Excel.WorkbookRangeAreas.ranges",
        description: "Returns ranges that comprise this object in a `RangeCollection` object.",
        kind: "Property",
        signature: "Excel.WorkbookRangeAreas.ranges: RangeCollection",
        examples: [],
      },
      {
        name: "Excel.WorkbookRangeAreas.getRangeAreasBySheet",
        description:
          "Returns the `RangeAreas` object based on worksheet ID or name in the collection.",
        kind: "Method",
        signature:
          "Excel.WorkbookRangeAreas.getRangeAreasBySheet => (key: string) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.WorkbookRangeAreas.getRangeAreasOrNullObjectBySheet",
        description:
          "Returns the `RangeAreas` object based on worksheet name or ID in the collection. If the worksheet does not exist, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.WorkbookRangeAreas.getRangeAreasOrNullObjectBySheet => (key: string) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.WorkbookRangeAreas.track",
        description:
          'Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created.',
        kind: "Method",
        signature: "Excel.WorkbookRangeAreas.track => () => Excel.WorkbookRangeAreas",
        examples: [],
      },
      {
        name: "Excel.WorkbookRangeAreas.untrack",
        description:
          "Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.",
        kind: "Method",
        signature: "Excel.WorkbookRangeAreas.untrack => () => Excel.WorkbookRangeAreas",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.Worksheet",
    apiList: [
      {
        name: "Excel.Worksheet.autoFilter",
        description: "Represents the `AutoFilter` object of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.autoFilter: Excel.AutoFilter",
        examples: [
          "activeWorksheet.autoFilter.clearColumnCriteria(3);",
          "activeWorksheet.autoFilter.reapply();",
          "activeWorksheet.autoFilter.remove();",
        ],
      },
      {
        name: "Excel.Worksheet.charts",
        description: "Returns a collection of charts that are part of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.charts: Excel.ChartCollection",
        examples: [
          "let chart = activeWorksheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.auto);",
          'const activeChart = activeWorksheet.charts.getItem("Chart1");',
          'let chart = activeWorksheet.charts.add(Excel.ChartType.columnStacked, activeWorksheet.getRange("B3:C5"));',
          'const activeChart = activeWorksheet.charts.getItem("SalesChart");',
          'const chart = workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");',
          'const activeChart = activeWorksheet.charts.getItem("Sales Chart");',
          "activeWorksheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);",
          "const charts = activeWorksheet.charts;",
          'const lastPosition = workbook.worksheets.getItem("Sheet1").charts.count - 1;',
          'const chart = workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);',
          'const activeChart = activeWorksheet.charts.getItem("Product Chart");',
          'let chart = activeWorksheet.charts.add("XYScatterSmooth", dataRange, "Auto");',
          "const bubbleChart = activeWorksheet.charts.add(Excel.ChartType.bubble, valueRange);",
          'let chart = sheet.charts.add("Line", dataRange, Excel.ChartSeriesBy.rows);',
          'let chart = activeWorksheet.charts.add(Excel.ChartType.line, dataRange, "Auto");',
        ],
      },
      {
        name: "Excel.Worksheet.comments",
        description: "Returns a collection of all the Comments objects on the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.comments: Excel.CommentCollection",
        examples: [
          "const comment = activeWorksheet.comments.getItemAt(0);",
          "activeWorksheet.comments.getItemAt(0).resolved = true;",
          'activeWorksheet.comments.add("A2", "TODO: add data.");',
          'activeWorksheet.comments.add("A1", commentBody, Excel.ContentType.mention);',
        ],
      },
      {
        name: "Excel.Worksheet.customProperties",
        description: "Gets a collection of worksheet-level custom properties.",
        kind: "Property",
        signature: "Excel.Worksheet.customProperties: WorksheetCustomPropertyCollection",
        examples: [],
      },
      {
        name: "Excel.Worksheet.enableCalculation",
        description:
          "Determines if Excel should recalculate the worksheet when necessary. True if Excel recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.",
        kind: "Property",
        signature: "Excel.Worksheet.enableCalculation: boolean",
        examples: [],
      },
      {
        name: "Excel.Worksheet.freezePanes",
        description: "Gets an object that can be used to manipulate frozen panes on the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.freezePanes: Excel.WorksheetFreezePanes",
        examples: [
          'activeWorksheet.freezePanes.freezeAt(activeWorksheet.getRange("H2:K5"));',
          "activeWorksheet.freezePanes.freezeColumns(2);",
          "activeWorksheet.freezePanes.freezeRows(2);",
          "const frozenRange = activeWorksheet.freezePanes.getLocationOrNullObject();",
          "activeWorksheet.freezePanes.unfreeze();",
        ],
      },
      {
        name: "Excel.Worksheet.horizontalPageBreaks",
        description:
          "Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks.",
        kind: "Property",
        signature: "Excel.Worksheet.horizontalPageBreaks: Excel.PageBreakCollection",
        examples: ['activeWorksheet.horizontalPageBreaks.add("A21:E21");'],
      },
      {
        name: "Excel.Worksheet.id",
        description:
          "Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.",
        kind: "Property",
        signature: "Excel.Worksheet.id: string",
        examples: [],
      },
      {
        name: "Excel.Worksheet.name",
        description: "The display name of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.name: string",
        examples: [
          '`The active worksheet is "${activeWorksheet.name}"`;',
          '`The name of the first worksheet is "${firstSheet.name}"`;',
          '`The name of the last worksheet is "${lastSheet.name}"`;',
          '`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`;',
          '`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`;',
          '`Added worksheet named "${sheet.name}" in position ${sheet.position}`;',
          'activeWorksheet.name = "New Name";',
          '`Worksheet with name "${activeWorksheet.name}" is hidden`;',
          '`Worksheet with name "${activeWorksheet.name}" is visible`;',
          '"\'" + activeWorksheet.name + "\' was copied to \'" + copiedSheet.name + "\'";',
          "let firstYear = firstSheet.name.substr(5, 4);",
          "let lastYear = lastSheet.name.substr(5, 4);",
          "let currentYear = activeWorksheet.name.substr(5, 4);",
          "let previousYear = previousYearSheet.name.substr(5, 4);",
          "worksheet.name;",
        ],
      },
      {
        name: "Excel.Worksheet.namedSheetViews",
        description: "Returns a collection of sheet views that are present in the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.namedSheetViews: NamedSheetViewCollection",
        examples: [],
      },
      {
        name: "Excel.Worksheet.names",
        description: "Collection of names scoped to the current worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.names: Excel.NamedItemCollection",
        examples: [
          'const myNamedItem = activeWorksheet.names.getItemOrNullObject("MyRange");',
          'activeWorksheet.names.add("ExpensesHeader", headerRange);',
        ],
      },
      {
        name: "Excel.Worksheet.optimization",
        description:
          "Returns a `WorksheetOptimization` that can scan and perform optimizations on the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.optimization: WorksheetOptimization",
        examples: [],
      },
      {
        name: "Excel.Worksheet.pageLayout",
        description: "Gets the `PageLayout` object of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.pageLayout: Excel.PageLayout",
        examples: [
          "activeWorksheet.pageLayout.centerHorizontally = true;",
          "activeWorksheet.pageLayout.centerVertically = true;",
          'activeWorksheet.pageLayout.setPrintTitleRows("$1:$1");',
          'activeWorksheet.pageLayout.setPrintArea("A1:D100");',
          "activeWorksheet.pageLayout.orientation = Excel.PageOrientation.landscape;",
          'activeWorksheet.pageLayout.setPrintArea("A1:D41");',
          "activeWorksheet.pageLayout.zoom = { scale: 200 };",
        ],
      },
      {
        name: "Excel.Worksheet.pivotTables",
        description: "Collection of PivotTables that are part of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.pivotTables: Excel.PivotTableCollection",
        examples: [
          'activeWorksheet.pivotTables.add("Farm Sales", "A1:E21", "A22");',
          'workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
          'const pivotTable = activeWorksheet.pivotTables.getItem("Farm Sales");',
          'const pivotTable = activeWorksheet.pivotTables.getItem("All Farm Sales");',
          'workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
        ],
      },
      {
        name: "Excel.Worksheet.position",
        description: "The zero-based position of the worksheet within the workbook.",
        kind: "Property",
        signature: "Excel.Worksheet.position: number",
        examples: [
          '`Added worksheet named "${sheet.name}" in position ${sheet.position}`;',
          "activeWorksheet.position = 2;",
        ],
      },
      {
        name: "Excel.Worksheet.protection",
        description: "Returns the sheet protection object for a worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.protection: Excel.WorksheetProtection",
        examples: ["activeWorksheet.protection.protect();"],
      },
      {
        name: "Excel.Worksheet.rangeValuesPreview",
        description:
          "Shows the preview of range values. Previews are non-persistent and have no co-authoring impact.",
        kind: "Property",
        signature: "Excel.Worksheet.rangeValuesPreview: RangeValuesPreview",
        examples: [],
      },
      {
        name: "Excel.Worksheet.shapes",
        description: "Returns the collection of all the Shape objects on the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.shapes: Excel.ShapeCollection",
        examples: [
          "let shapes = activeWorksheet.shapes;",
          "const shapes = activeWorksheet.shapes;",
          "const shape = activeWorksheet.shapes.addGeometricShape(Excel.GeometricShapeType.hexagon);",
          'const image = activeWorksheet.shapes.getItem("Image").image;',
          'const shape = activeWorksheet.shapes.getItem("Image");',
          "const shapes = sheet.shapes;",
          "const shape = activeWorksheet.shapes.addGeometricShape(Excel.GeometricShapeType.smileyFace);",
          'const shapeGroup = activeWorksheet.shapes.getItem("Group").group;',
          'const shape = activeWorksheet.shapes.getItem("Square");',
          'const shape = activeWorksheet.shapes.getItem("Pentagon");',
          'const shape = activeWorksheet.shapes.getItem("Octagon");',
          "const shape = activeWorksheet.shapes.addGeometricShape(Excel.GeometricShapeType.triangle);",
          'const square = activeWorksheet.shapes.getItem("Square");',
          'const pentagon = activeWorksheet.shapes.getItem("Pentagon");',
          'const octagon = activeWorksheet.shapes.getItem("Octagon");',
          "const shapeGroup = activeWorksheet.shapes.addGroup([square, pentagon, octagon]);",
        ],
      },
      {
        name: "Excel.Worksheet.showGridlines",
        description: "Specifies if gridlines are visible to the user.",
        kind: "Property",
        signature: "Excel.Worksheet.showGridlines: boolean",
        examples: ["activeWorksheet.showGridlines = true;"],
      },
      {
        name: "Excel.Worksheet.showHeadings",
        description: "Specifies if headings are visible to the user.",
        kind: "Property",
        signature: "Excel.Worksheet.showHeadings: boolean",
        examples: [],
      },
      {
        name: "Excel.Worksheet.slicers",
        description: "Returns a collection of slicers that are part of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.slicers: Excel.SlicerCollection",
        examples: [
          'let slicer = activeWorksheet.slicers.add("Farm Sales", "Type");',
          "activeWorksheet.slicers.getItemAt(0).delete();",
          'const slicer = activeWorksheet.slicers.add("Farm Sales", "Type");',
        ],
      },
      {
        name: "Excel.Worksheet.standardHeight",
        description:
          "Returns the standard (default) height of all the rows in the worksheet, in points.",
        kind: "Property",
        signature: "Excel.Worksheet.standardHeight: number",
        examples: [],
      },
      {
        name: "Excel.Worksheet.standardWidth",
        description:
          "Specifies the standard (default) width of all the columns in the worksheet. One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.",
        kind: "Property",
        signature: "Excel.Worksheet.standardWidth: number",
        examples: [],
      },
      {
        name: "Excel.Worksheet.tabColor",
        description:
          'The tab color of the worksheet. When retrieving the tab color, if the worksheet is invisible, the value will be `null`. If the worksheet is visible but the tab color is set to auto, an empty string will be returned. Otherwise, the property will be set to a color, in the form #RRGGBB (e.g., "FFA500"). When setting the color, use an empty-string to set an "auto" color, or a real color otherwise.',
        kind: "Property",
        signature: "Excel.Worksheet.tabColor: string",
        examples: ['activeWorksheet.tabColor = "#FF0000";'],
      },
      {
        name: "Excel.Worksheet.tabId",
        description:
          'Returns a value representing this worksheet that can be read by Open Office XML. This is an integer value, which is different from `worksheet.id` (which returns a globally unique identifier) and `worksheet.name` (which returns a value such as "Sheet1").',
        kind: "Property",
        signature: "Excel.Worksheet.tabId: number",
        examples: [],
      },
      {
        name: "Excel.Worksheet.tables",
        description: "Collection of tables that are part of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.tables: Excel.TableCollection",
        examples: [
          'const activeTable = activeWorksheet.tables.getItem("TemperatureTable");',
          'activeWorksheet.tables.add("B2:E5", true);',
          'const activeTable = activeWorksheet.tables.getItem("AthletesTable");',
          'const activeTable = activeWorksheet.tables.getItem("ExpensesTable");',
          'let expensesTable = activeWorksheet.tables.add("A1:D1", true);',
          'let expensesTable = activeWorksheet.tables.add("A1:E7", true);',
          'let table = activeWorksheet.tables.add("A1:B3", true);',
          'const activeTable = activeWorksheet.tables.getItem("SalesTable");',
          'const activeTable = activeWorksheet.tables.getItem("Sales");',
          'let expensesTable = sheet.tables.add("A1:E1", true);',
          'const activeTable = activeWorksheet.tables.getItem("Table1");',
          'const activeTable = activeWorksheet.tables.getItem("NameOptionsTable");',
          'const activeTable = activeWorksheet.tables.getItem("Table2");',
          'const activeTable = activeWorksheet.tables.getItem("Table5");',
          'const activeTable = activeWorksheet.tables.getItem("ProductSales");',
          'const activeTable = activeWorksheet.tables.getItem("UnfilteredTable");',
          'const newTable = activeWorksheet.tables.add("G1:K1", true);',
          'const newTable = activeWorksheet.tables.add("G1:J1", true);',
        ],
      },
      {
        name: "Excel.Worksheet.tasks",
        description: "Returns a collection of tasks that are present in the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.tasks: DocumentTaskCollection",
        examples: [],
      },
      {
        name: "Excel.Worksheet.verticalPageBreaks",
        description:
          "Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks.",
        kind: "Property",
        signature: "Excel.Worksheet.verticalPageBreaks: PageBreakCollection",
        examples: [],
      },
      {
        name: "Excel.Worksheet.visibility",
        description: "The visibility of the worksheet.",
        kind: "Property",
        signature:
          'Excel.Worksheet.visibility: Excel.SheetVisibility | "Visible" | "Hidden" | "VeryHidden"',
        examples: [
          "activeWorksheet.visibility = Excel.SheetVisibility.hidden;",
          "activeWorksheet.visibility = Excel.SheetVisibility.visible;",
        ],
      },
      {
        name: "Excel.Worksheet.visuals",
        description: "Returns a collection of visuals that are part of the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.visuals: VisualCollection",
        examples: [],
      },
      {
        name: "Excel.Worksheet.visualTracker",
        description: "Returns the visual tracker associated with the worksheet.",
        kind: "Property",
        signature: "Excel.Worksheet.visualTracker: VisualTracker",
        examples: [],
      },
      {
        name: "Excel.Worksheet.activate",
        description: "Activate the worksheet in the Excel UI.",
        kind: "Method",
        signature: "Excel.Worksheet.activate() => void",
        examples: ["activeWorksheet.activate();", "sheet.activate();"],
      },
      {
        name: "Excel.Worksheet.calculate",
        description: "Calculates all cells on a worksheet.",
        kind: "Method",
        signature: "Excel.Worksheet.calculate => (markAllDirty: boolean) => void",
        examples: [],
      },
      {
        name: "Excel.Worksheet.copy",
        description: "Copies a worksheet and places it at the specified position.",
        kind: "Method",
        signature:
          "Excel.Worksheet.copy(positionType?: Excel.WorksheetPositionType, relativeTo?: Excel.Worksheet): Excel.Worksheet",
        examples: [
          "activeWorksheet.copy(Excel.WorksheetPositionType.after, activeWorksheet);",
          'let copiedSheet = activeWorksheet.copy("End");',
        ],
      },
      {
        name: "Excel.Worksheet.delete",
        description:
          'Deletes the worksheet from the workbook. Note that if the worksheet\'s visibility is set to "VeryHidden", the delete operation will fail with an `InvalidOperation` exception. You should first change its visibility to hidden or visible before deleting it.',
        kind: "Method",
        signature: "Excel.Worksheet.delete() => void",
        examples: [
          'workbook.worksheets.getItemOrNullObject("Sample").delete();',
          'workbook.worksheets.getItemOrNullObject("Shapes").delete();',
          "activeWorksheet.delete();",
        ],
      },
      {
        name: "Excel.Worksheet.findAll",
        description:
          "Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.",
        kind: "Method",
        signature:
          "Excel.Worksheet.findAll(text: string, criteria: Excel.WorksheetSearchCriteria) => Excel.RangeAreas",
        examples: [
          'let foundRanges = activeWorksheet.findAll("Complete", {\n    completeMatch: true,\n    matchCase: false,\n  });',
        ],
      },
      {
        name: "Excel.Worksheet.findAllOrNullObject",
        description:
          "Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.",
        kind: "Method",
        signature:
          "Excel.Worksheet.findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getCell",
        description:
          "Gets the `Range` object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.",
        kind: "Method",
        signature: "Excel.Worksheet.getCell(row: number, column: number) => Excel.Range",
        examples: [
          "let cell = activeWorksheet.getCell(1, 4);",
          "const cell = activeWorksheet.getCell(0, 0);",
        ],
      },
      {
        name: "Excel.Worksheet.getNext",
        description:
          "Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.",
        kind: "Method",
        signature: "Excel.Worksheet.getNext(visibleOnly?: boolean) => Excel.Worksheet",
        examples: [
          "let nextSheet = activeWorksheet.getNext();",
          "const firstSheet = sheets.getFirst().getNext();",
        ],
      },
      {
        name: "Excel.Worksheet.getNextOrNullObject",
        description:
          "Gets the worksheet that follows this one. If there are no worksheets following this one, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.Worksheet.getNextOrNullObject => (visibleOnly?: boolean) => Excel.Worksheet",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getPrevious",
        description:
          "Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.",
        kind: "Method",
        signature: "Excel.Worksheet.getPrevious(visibleOnly?: boolean) => Excel.Worksheet",
        examples: [
          "let previousSheet = activeWorksheet.getPrevious();",
          "const previousYearSheet = activeWorksheet.getPrevious();",
        ],
      },
      {
        name: "Excel.Worksheet.getPreviousOrNullObject",
        description:
          "Gets the worksheet that precedes this one. If there are no previous worksheets, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.Worksheet.getPreviousOrNullObject => (visibleOnly?: boolean) => Excel.Worksheet",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getRange",
        description:
          "Gets the `Range` object, representing a single rectangular block of cells, specified by the address or name.",
        kind: "Method",
        signature: "Excel.Worksheet.getRange(address?: string) => Excel.Range",
        examples: [
          'let dataRange = activeWorksheet.getRange("A1:B13");',
          'let dataRange = activeWorksheet.getRange("D2:D5");',
          'const range = activeWorksheet.getRange("B21:E23");',
          'const range = activeWorksheet.getRange("B2:M5");',
          'const range = activeWorksheet.getRange("B8:E13");',
          'const range = activeWorksheet.getRange("B16:D18");',
          "const range = activeWorksheet.getRange();",
          'let headerRange = activeWorksheet.getRange("B2:E2");',
          'let dataRange = activeWorksheet.getRange("B3:D5");',
          'let totalRange = activeWorksheet.getRange("E3:E6");',
          'let chart = activeWorksheet.charts.add(Excel.ChartType.columnStacked, activeWorksheet.getRange("B3:C5"));',
          'let range = activeWorksheet.getRange("B2:C5");',
          'let pinkColumnRange = activeWorksheet.getRange("H:H");',
          'let rangeToAnalyze = workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");',
          'let rangeToPlacePivot = workbook.worksheets.getItem("PivotWorksheet").getRange("A2");',
          'let masterTotalRange = activeWorksheet.getRange("E30");',
          'let range = activeWorksheet.getRange("E2:E5");',
          'let range = activeWorksheet.getRange("B4:E4");',
          'activeWorksheet.getRange("G1").copyFrom("A1:E1");',
          'activeWorksheet.getRange("D1").copyFrom("A1:C1", Excel.RangeCopyType.all, true, false);',
          'activeWorksheet.getRange("D2").copyFrom("A2:C2", Excel.RangeCopyType.all, false, false);',
          'activeWorksheet.getRange("F1").values = [["Moved Range"]];',
          'activeWorksheet.getRange("A1:E1").moveTo("G1");',
          'let targetCell = activeWorksheet.getRange("G4");',
          'let range = activeWorksheet.getRange("MyRange");',
          "let range = activeWorksheet.getRange();",
          'activeWorksheet.getRange("4:9").group(Excel.GroupOption.byRows);',
          'activeWorksheet.getRange("4:5").group(Excel.GroupOption.byRows);',
          'activeWorksheet.getRange("7:8").group(Excel.GroupOption.byRows);',
          'activeWorksheet.getRange("C:Q").group(Excel.GroupOption.byColumns);',
          'activeWorksheet.getRange("C:F").group(Excel.GroupOption.byColumns);',
          'activeWorksheet.getRange("H:K").group(Excel.GroupOption.byColumns);',
          'activeWorksheet.getRange("M:P").group(Excel.GroupOption.byColumns);',
          'let range = activeWorksheet.getRange("B2:D11");',
          'let range = activeWorksheet.getRange("B2:E2");',
          'let range = activeWorksheet.getRange("D3:E5");',
          'let range = activeWorksheet.getRange("C3");',
          'let range = activeWorksheet.getRange("B5:D5");',
          'let range = activeWorksheet.getRange("E3");',
          'let range = activeWorksheet.getRange("E3:E6");',
          'let range = activeWorksheet.getRange("B2:E6");',
          'activeWorksheet.getRange("A11:A11").values = [["Results"]];',
          'activeWorksheet.getRange("A13:D13").values = headerValues;',
          'activeWorksheet.getRange("A14:D20").values = bodyValues;',
          'activeWorksheet.getRange("B23:B29").values = merchantColumnValues;',
          'activeWorksheet.getRange("A32:D32").values = secondRowValues;',
          'let range = activeWorksheet.getRange("A1:E7");',
          'let range = activeWorksheet.getRange("A1:D4");',
          'rangeToSet = activeWorksheet.getRange("A1:C1");',
          'rangeToGet = activeWorksheet.getRange("A1:C1");',
          'rangeToSet = activeWorksheet.getRange("A1:B1");',
          'let range = activeWorksheet.getRange("A1:B3");',
          'const sumCell = activeWorksheet.getRange("K4");',
          'const range = activeWorksheet.getRange("A1:E5");',
          'const range = sheet.getRange("A1");',
          'const sourceData = activeWorksheet.getRange("A1:B4");',
          "const range = workbook.worksheets.getItem(sheetName).getRange(rangeSelection);",
          "const range = activeWorksheet.getRange(rangeSelection);",
          'let rangeSelection = activeWorksheet.getRange("C2:C7");',
          'let xRangeSelection = activeWorksheet.getRange("A1:A7");',
          'let dataRange = sheet.getRange("A1:E7");',
          'let dataRange = activeWorksheet.getRange("A1:E7");',
          'const productsRange = activeWorksheet.getRange("A3:A11");',
          "const range = activeWorksheet.getRange(rangeAddress);",
          'const dateTimeData = activeWorksheet.getRange("A2:B6");',
          'const range = activeWorksheet.getRange("A1:A5");',
          'const nameSourceRange = workbook.worksheets.getItem("Names").getRange("A1:A3");',
          'const range = activeWorksheet.getRange("A5:F5");',
          'const currencyRange = sheet.getRange("A2");',
          'const dateRange = sheet.getRange("A1");',
          'const tableRange = activeWorksheet.getRange("B2:E6");',
          'const range = activeWorksheet.getRange("B4:E4");',
          'activeWorksheet.getRange("B10:D14").select();',
          'const headerRange = activeWorksheet.getRange("A1:D1");',
          'const bigNumberSource = activeWorksheet.getRange("B3");',
          'const resultRange = activeWorksheet.getRange("C3");',
          'const masterTotalRange = activeWorksheet.getRange("B27:C27");',
          'const rangeToAnalyze = workbook.worksheets.getItem("Data").getRange("A1:E21");',
          'const rangeToPlacePivot = workbook.worksheets.getItem("Pivot").getRange("A2");',
          'const sumCell = activeWorksheet.getRange("P4");',
          'activeWorksheet.getRange("F2").values = [["Copied Formula"]];',
          'activeWorksheet.getRange("G2").copyFrom("A1:E1", Excel.RangeCopyType.formulas);',
          "let range = activeWorksheet.getRange(rangeAddress);",
          "const range = activeWorksheet.getRange(rangeAddress).getColumn(1);",
          'const range = activeWorksheet.getRange(rangeAddress).getIntersection("D4:G6");',
          "const range = activeWorksheet.getRange(rangeAddress).getLastCell();",
          "const range = activeWorksheet.getRange(rangeAddress).getLastColumn();",
          "const range = activeWorksheet.getRange(rangeAddress).getLastRow();",
          "const range = activeWorksheet.getRange(rangeAddress).getOffsetRange(-1, 4);",
          "const range = activeWorksheet.getRange(rangeAddress).getRow(1);",
          'const targetCell = activeWorksheet.getRange("G4");',
          'let productsRange = activeWorksheet.getRange("A3:A5");',
          'activeWorksheet.getRange("F12").values = [["Moved Range:"]];',
          'activeWorksheet.getRange("A1:E1").moveTo("G12");',
          'const range = activeWorksheet.getRange("B2:D11");',
          'const sourceRange = activeWorksheet.getRange("B2:E2");',
          'const targetRange = activeWorksheet.getRange("B7:E7");',
          'let range = activeWorksheet.getRange("A1:E1");',
          'const range = activeWorksheet.getRange("B2:E2");',
          'let productsRange = activeWorksheet.getRange("A9:A11");',
          'const firstTaxRateRange = firstSheet.getRange("B2");',
          'const lastTaxRateRange = lastSheet.getRange("B2");',
          'const currentTaxDueRange = activeWorksheet.getRange("C2");',
          'const previousTaxDueRange = previousYearSheet.getRange("C2");',
          'activeWorksheet.freezePanes.freezeAt(activeWorksheet.getRange("H2:K5"));',
        ],
      },
      {
        name: "Excel.Worksheet.getRangeByIndexes",
        description:
          "Gets the `Range` object beginning at a particular row index and column index, and spanning a certain number of rows and columns.",
        kind: "Method",
        signature:
          "Excel.Worksheet.getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number) => Excel.Range",
        examples: [
          "const pasteToRange = activeWorksheet.getRangeByIndexes(\n    0,\n    usedRange.columnCount + 1,\n    expensesTableValues.length,\n    expensesTableValues[0].length\n  );",
        ],
      },
      {
        name: "Excel.Worksheet.getRangeR1C1",
        description:
          "Gets the `Range` object, representing a single rectangular block of cells, specified by the address in R1C1 format.",
        kind: "Method",
        signature: "Excel.Worksheet.getRangeR1C1 => (address: string) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getRanges",
        description:
          "Gets the `RangeAreas` object, representing one or more blocks of rectangular ranges, specified by the address or name.",
        kind: "Method",
        signature: "Excel.Worksheet.getRanges(address?: string) => Excel.RangeAreas",
        examples: [
          'let rangeAreas = activeWorksheet.getRanges("F3:F5, H3:H5");',
          'let rangeAreas = activeWorksheet.getRanges("F:F, H:H");',
          'let rangeAreas = activeWorksheet.getRanges("F3:F5, H:H");',
          'const specifiedRanges = activeWorksheet.getRanges("D3:D5, G3:G5");',
        ],
      },
      {
        name: "Excel.Worksheet.getRangesR1C1",
        description:
          "Gets the `RangeAreas` object, representing one or more blocks of rectangular ranges, specified by the address in R1C1 format.",
        kind: "Method",
        signature: "Excel.Worksheet.getRangesR1C1 => (address: string) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getUsedRange",
        description:
          "The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e. it will *not* throw an error).",
        kind: "Method",
        signature: "Excel.Worksheet.getUsedRange(valuesOnly?: boolean) => Excel.Range",
        examples: [
          "let range = activeWorksheet.getUsedRange();",
          "let usedRange = activeWorksheet.getUsedRange();",
          "activeWorksheet.getUsedRange().format.autofitColumns();",
          "activeWorksheet.getUsedRange().format.autofitRows();",
          "const farmData = activeWorksheet.getUsedRange();",
          "sheet.getUsedRange().format.autofitColumns();",
          "sheet.getUsedRange().format.autofitRows();",
          "const usedRange = activeWorksheet.getUsedRange();",
        ],
      },
      {
        name: "Excel.Worksheet.getUsedRangeAreas",
        description:
          'Returns a set of rectangular regions of data in the worksheet. Each region is an "island" of contiguous data.',
        kind: "Method",
        signature:
          "Excel.Worksheet.getUsedRangeAreas => (options?: Excel.GetUsedRangeAreasOptions) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getUsedRangeAreasOrNullObject",
        description:
          'Returns a set of rectangular regions of data in the worksheet. Each region is an "island" of contiguous data. If there are no regions of data, then this function will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.',
        kind: "Method",
        signature:
          "Excel.Worksheet.getUsedRangeAreasOrNullObject => (options?: Excel.GetUsedRangeAreasOptions) => Excel.RangeAreas",
        examples: [],
      },
      {
        name: "Excel.Worksheet.getUsedRangeOrNullObject",
        description:
          "The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature:
          "Excel.Worksheet.getUsedRangeOrNullObject => (valuesOnly?: boolean) => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.Worksheet.replaceAll",
        description:
          "Finds and replaces the given string based on the criteria specified within the current worksheet.",
        kind: "Method",
        signature:
          "Excel.Worksheet.replaceAll => (text: string, replacement: string, criteria: Excel.ReplaceCriteria) => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.Worksheet.showOutlineLevels",
        description:
          "Shows row or column groups by their outline levels. Outlines groups and summarizes a list of data in the worksheet. The `rowLevels` and `columnLevels` parameters specify how many levels of the outline will be displayed. The acceptable argument range is between 0 and 8. A value of 0 does not change the current display. A value greater than the current number of levels displays all the levels.",
        kind: "Method",
        signature:
          "Excel.Worksheet.showOutlineLevels => (rowLevels: number, columnLevels: number) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetCollection",
    apiList: [
      {
        name: "Excel.WorksheetCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.WorksheetCollection.items: Excel.Worksheet[]",
        examples: [],
      },
      {
        name: "Excel.WorksheetCollection.add",
        description:
          "Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call `.activate()` on it.",
        kind: "Method",
        signature: "Excel.WorksheetCollection.add(name?: string) => Excel.Worksheet",
        examples: [
          'let sheet = sheets.add("Sample");',
          'const sheet = workbook.worksheets.add("Sample");',
          'const sheet = workbook.worksheets.add("Shapes");',
          "const worksheet = workbook.worksheets.add(wSheetName);",
        ],
      },
      {
        name: "Excel.WorksheetCollection.getActiveWorksheet",
        description: "Gets the currently active worksheet in the workbook.",
        kind: "Method",
        signature: "Excel.WorksheetCollection.getActiveWorksheet() => Excel.Worksheet",
        examples: ["const activeWorksheet = workbook.worksheets.getActiveWorksheet();"],
      },
      {
        name: "Excel.WorksheetCollection.getCount",
        description: "Gets the number of worksheets in the collection.",
        kind: "Method",
        signature:
          "Excel.WorksheetCollection.getCount => (visibleOnly?: boolean) => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.WorksheetCollection.getFirst",
        description: "Gets the first worksheet in the collection.",
        kind: "Method",
        signature: "Excel.WorksheetCollection.getFirst(visibleOnly?: boolean) => Excel.Worksheet",
        examples: [
          "let firstSheet = workbook.worksheets.getFirst();",
          "const firstSheet = sheets.getFirst().getNext();",
        ],
      },
      {
        name: "Excel.WorksheetCollection.getItem",
        description: "Gets a worksheet object using its name or ID.",
        kind: "Method",
        signature: "Excel.WorksheetCollection.getItem(key: string) => Excel.Worksheet",
        examples: [
          'let rangeToAnalyze = workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");',
          'let rangeToPlacePivot = workbook.worksheets.getItem("PivotWorksheet").getRange("A2");',
          'workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
          "const range = workbook.worksheets.getItem(sheetName).getRange(rangeSelection);",
          'const chart = workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");',
          'const lastPosition = workbook.worksheets.getItem("Sheet1").charts.count - 1;',
          'const chart = workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);',
          'const nameSourceRange = workbook.worksheets.getItem("Names").getRange("A1:A3");',
          'const rangeToAnalyze = workbook.worksheets.getItem("Data").getRange("A1:E21");',
          'const rangeToPlacePivot = workbook.worksheets.getItem("Pivot").getRange("A2");',
          'workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);',
        ],
      },
      {
        name: "Excel.WorksheetCollection.getLast",
        description: "Gets the last worksheet in the collection.",
        kind: "Method",
        signature: "Excel.WorksheetCollection.getLast(visibleOnly?: boolean) => Excel.Worksheet",
        examples: [
          "let lastSheet = workbook.worksheets.getLast();",
          "const lastSheet = sheets.getLast();",
        ],
      },
    ],
  },
  {
    objName: "Excel.WorksheetCustomProperty",
    apiList: [
      {
        name: "Excel.WorksheetCustomProperty.key",
        description:
          "Gets the key of the custom property. Custom property keys are case-insensitive. The key is limited to 255 characters (larger values will cause an `InvalidArgument` error to be thrown.)",
        kind: "Property",
        signature: "Excel.WorksheetCustomProperty.key: string",
        examples: [],
      },
      {
        name: "Excel.WorksheetCustomProperty.value",
        description: "Gets or sets the value of the custom property.",
        kind: "Property",
        signature: "Excel.WorksheetCustomProperty.value: string",
        examples: [],
      },
      {
        name: "Excel.WorksheetCustomProperty.delete",
        description: "Deletes the custom property.",
        kind: "Method",
        signature: "Excel.WorksheetCustomProperty.delete => () => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetCustomPropertyCollection",
    apiList: [
      {
        name: "Excel.WorksheetCustomPropertyCollection.items",
        description: "Gets the loaded child items in this collection.",
        kind: "Property",
        signature: "Excel.WorksheetCustomPropertyCollection.items: WorksheetCustomProperty[]",
        examples: [],
      },
      {
        name: "Excel.WorksheetCustomPropertyCollection.add",
        description:
          "Adds a new custom property that maps to the provided key. This overwrites existing custom properties with that key.",
        kind: "Method",
        signature:
          "Excel.WorksheetCustomPropertyCollection.add => (key: string, value: string) => Excel.WorksheetCustomProperty",
        examples: [],
      },
      {
        name: "Excel.WorksheetCustomPropertyCollection.getCount",
        description: "Gets the number of custom properties on this worksheet.",
        kind: "Method",
        signature:
          "Excel.WorksheetCustomPropertyCollection.getCount => () => OfficeExtension.ClientResult<number>",
        examples: [],
      },
      {
        name: "Excel.WorksheetCustomPropertyCollection.getItem",
        description:
          "Gets a custom property object by its key, which is case-insensitive. Throws an error if the custom property does not exist.",
        kind: "Method",
        signature:
          "Excel.WorksheetCustomPropertyCollection.getItem => (key: string) => Excel.WorksheetCustomProperty",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetFreezePanes",
    apiList: [
      {
        name: "Excel.WorksheetFreezePanes.freezeAt",
        description:
          "Sets the frozen cells in the active worksheet view. The range provided corresponds to cells that will be frozen in the top- and left-most pane.",
        kind: "Method",
        signature: "Excel.WorksheetFreezePanes.freezeAt(frozenRange: string | Excel.Range) => void",
        examples: ['activeWorksheet.freezePanes.freezeAt(activeWorksheet.getRange("H2:K5"));'],
      },
      {
        name: "Excel.WorksheetFreezePanes.freezeColumns",
        description: "Freeze the first column or columns of the worksheet in place.",
        kind: "Method",
        signature: "Excel.WorksheetFreezePanes.freezeColumns(count?: number) => void",
        examples: ["activeWorksheet.freezePanes.freezeColumns(2);"],
      },
      {
        name: "Excel.WorksheetFreezePanes.freezeRows",
        description: "Freeze the top row or rows of the worksheet in place.",
        kind: "Method",
        signature: "Excel.WorksheetFreezePanes.freezeRows(count?: number) => void",
        examples: ["activeWorksheet.freezePanes.freezeRows(2);"],
      },
      {
        name: "Excel.WorksheetFreezePanes.getLocation",
        description:
          "Gets a range that describes the frozen cells in the active worksheet view. The frozen range corresponds to cells that are frozen in the top- and left-most pane.",
        kind: "Method",
        signature: "Excel.WorksheetFreezePanes.getLocation => () => Excel.Range",
        examples: [],
      },
      {
        name: "Excel.WorksheetFreezePanes.getLocationOrNullObject",
        description:
          "Gets a range that describes the frozen cells in the active worksheet view. The frozen range corresponds to cells that are frozen in the top- and left-most pane. If there is no frozen pane, then this method returns an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.",
        kind: "Method",
        signature: "Excel.WorksheetFreezePanes.getLocationOrNullObject() => Excel.Range",
        examples: ["const frozenRange = activeWorksheet.freezePanes.getLocationOrNullObject();"],
      },
      {
        name: "Excel.WorksheetFreezePanes.unfreeze",
        description: "Removes all frozen panes in the worksheet.",
        kind: "Method",
        signature: "Excel.WorksheetFreezePanes.unfreeze() => void",
        examples: ["activeWorksheet.freezePanes.unfreeze();"],
      },
    ],
  },
  {
    objName: "Excel.WorksheetOptimization",
    apiList: [
      {
        name: "Excel.WorksheetOptimization.optimize",
        description:
          "Optimizes the worksheet, returning the number of cells that were allocated and the number of cells that were optimized.",
        kind: "Method",
        signature:
          "Excel.WorksheetOptimization.optimize => () => Excel.WorksheetOptimizationResult",
        examples: [],
      },
      {
        name: "Excel.WorksheetOptimization.scan",
        description:
          "Scans the worksheet for optimizations that can be made, returning a collection of potential optimizations.",
        kind: "Method",
        signature: "Excel.WorksheetOptimization.scan => () => Excel.RangeOptimizationCollection",
        examples: [],
      },
      {
        name: "Excel.WorksheetOptimization.scanExtended",
        description:
          "Scan the worksheet for optimizations that can be made, returning allocatedCells, optimizableCells, and the collection of optimizations that can be made. This is created to replace the original scan() to give the option to extend additional types of optimizable content, and to avoid the expensive enumeration of entire collection to request the cell properties.",
        kind: "Method",
        signature:
          "Excel.WorksheetOptimization.scanExtended => () => Excel.WorksheetOptimizationScanResult",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetOptimizationResult",
    apiList: [
      {
        name: "Excel.WorksheetOptimizationResult.allocatedCells",
        description:
          "The number of cells that were allocated in the worksheet before the optimization took place.",
        kind: "Property",
        signature: "Excel.WorksheetOptimizationResult.allocatedCells: number",
        examples: [],
      },
      {
        name: "Excel.WorksheetOptimizationResult.optimizedCells",
        description: "The number of cells that were optimized.",
        kind: "Property",
        signature: "Excel.WorksheetOptimizationResult.optimizedCells: number",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetOptimizationScanResult",
    apiList: [
      {
        name: "Excel.WorksheetOptimizationScanResult.allocatedCells",
        description: "The number of cells that are allocated in the worksheet.",
        kind: "Property",
        signature: "Excel.WorksheetOptimizationScanResult.allocatedCells: number",
        examples: [],
      },
      {
        name: "Excel.WorksheetOptimizationScanResult.optimizableCells",
        description: "The number of cells in the worksheet that can be optimized.",
        kind: "Property",
        signature: "Excel.WorksheetOptimizationScanResult.optimizableCells: number",
        examples: [],
      },
      {
        name: "Excel.WorksheetOptimizationScanResult.ranges",
        description: "The collection of ranges that can be optimized.",
        kind: "Property",
        signature: "Excel.WorksheetOptimizationScanResult.ranges: RangeOptimizationCollection",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetProtection",
    apiList: [
      {
        name: "Excel.WorksheetProtection.allowEditRanges",
        description:
          "Specifies the `AllowEditRangeCollection` object found in this worksheet. This is a collection of `AllowEditRange` objects, which work with worksheet protection properties. When worksheet protection is enabled, an `AllowEditRange` object can be used to allow editing of a specific range, while maintaining protection on the rest of the worksheet.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.allowEditRanges: AllowEditRangeCollection",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.canPauseProtection",
        description: "Specifies if protection can be paused for this worksheet.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.canPauseProtection: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.isPasswordProtected",
        description: "Specifies if the sheet is password protected.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.isPasswordProtected: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.isPaused",
        description: "Specifies if worksheet protection is paused.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.isPaused: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.options",
        description: "Specifies the protection options for the worksheet.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.options: WorksheetProtectionOptions",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.protected",
        description: "Specifies if the worksheet is protected.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.protected: boolean",
        examples: [
          "if (!activeWorksheet.protection.protected) {\n    activeWorksheet.protection.protect();\n  }",
        ],
      },
      {
        name: "Excel.WorksheetProtection.savedOptions",
        description:
          "Specifies the protection options saved in the worksheet. This will return the same `WorksheetProtectionOptions` object regardless of the worksheet protection state.",
        kind: "Property",
        signature: "Excel.WorksheetProtection.savedOptions: WorksheetProtectionOptions",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.checkPassword",
        description:
          "Specifies if the password can be used to unlock worksheet protection. This method doesn't change the worksheet protection state. If a password is input but no password is required to unlock worksheet protection, this method will return false.",
        kind: "Method",
        signature:
          "Excel.WorksheetProtection.checkPassword => (password?: string) => OfficeExtension.ClientResult<boolean>",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.pauseProtection",
        description:
          "Pauses worksheet protection for the given worksheet object for the user in the current session. This method does nothing if worksheet protection isn't enabled or is already paused. If the password is incorrect, then this method throws an `InvalidArgument` error and fails to pause protection. This method does not change the protection state if worksheet protection is not enabled or already paused.",
        kind: "Method",
        signature: "Excel.WorksheetProtection.pauseProtection => (password?: string) => void",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.protect",
        description: "Protects a worksheet. Fails if the worksheet has already been protected.",
        kind: "Method",
        signature:
          "Excel.WorksheetProtection.protect(options?: Excel.WorksheetProtectionOptions, password?: string) => void",
        examples: ["activeWorksheet.protection.protect();"],
      },
      {
        name: "Excel.WorksheetProtection.resumeProtection",
        description:
          "Resumes worksheet protection for the given worksheet object for the user in a given session. Worksheet protection must be paused for this method to work. If worksheet protection is not paused, then this method will not change the protection state of the worksheet.",
        kind: "Method",
        signature: "Excel.WorksheetProtection.resumeProtection => () => void",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.unprotect",
        description: "Unprotects a worksheet.",
        kind: "Method",
        signature: "Excel.WorksheetProtection.unprotect => (password?: string) => void",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtection.updateOptions",
        description:
          "Change the worksheet protection options associated with the `WorksheetProtection` object. Worksheet protection must be disabled or paused for this method to work properly. If worksheet protection is enabled and not paused, this method throws an `AccessDenied` error and fails to change the worksheet protection options.",
        kind: "Method",
        signature:
          "Excel.WorksheetProtection.updateOptions => (options: Excel.WorksheetProtectionOptions) => void",
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetProtectionOptions",
    apiList: [
      {
        name: "Excel.WorksheetProtectionOptions.allowAutoFilter",
        description:
          "Represents the worksheet protection option allowing use of the AutoFilter feature.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowAutoFilter: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowDeleteColumns",
        description: "Represents the worksheet protection option allowing deleting of columns.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowDeleteColumns: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowDeleteRows",
        description: "Represents the worksheet protection option allowing deleting of rows.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowDeleteRows: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowEditObjects",
        description: "Represents the worksheet protection option allowing editing of objects.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowEditObjects: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowEditScenarios",
        description: "Represents the worksheet protection option allowing editing of scenarios.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowEditScenarios: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowFormatCells",
        description: "Represents the worksheet protection option allowing formatting of cells.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowFormatCells: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowFormatColumns",
        description: "Represents the worksheet protection option allowing formatting of columns.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowFormatColumns: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowFormatRows",
        description: "Represents the worksheet protection option allowing formatting of rows.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowFormatRows: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowInsertColumns",
        description: "Represents the worksheet protection option allowing inserting of columns.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowInsertColumns: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowInsertHyperlinks",
        description: "Represents the worksheet protection option allowing inserting of hyperlinks.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowInsertHyperlinks: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowInsertRows",
        description: "Represents the worksheet protection option allowing inserting of rows.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowInsertRows: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowPivotTables",
        description:
          "Represents the worksheet protection option allowing use of the PivotTable feature.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowPivotTables: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.allowSort",
        description: "Represents the worksheet protection option allowing use of the sort feature.",
        kind: "Property",
        signature: "Excel.WorksheetProtectionOptions.allowSort: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetProtectionOptions.selectionMode",
        description: "Represents the worksheet protection option of selection mode.",
        kind: "Property",
        signature:
          'Excel.WorksheetProtectionOptions.selectionMode: "None" | ProtectionSelectionMode | "Normal" | "Unlocked"',
        examples: [],
      },
    ],
  },
  {
    objName: "Excel.WorksheetSearchCriteria",
    apiList: [
      {
        name: "Excel.WorksheetSearchCriteria.completeMatch",
        description:
          "Specifies if the match needs to be complete or partial. A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`). Default is `false` (partial).",
        kind: "Property",
        signature: "Excel.WorksheetSearchCriteria.completeMatch: boolean",
        examples: [],
      },
      {
        name: "Excel.WorksheetSearchCriteria.matchCase",
        description:
          "Specifies if the match is case-sensitive. Default is `false` (case-insensitive).",
        kind: "Property",
        signature: "Excel.WorksheetSearchCriteria.matchCase: boolean",
        examples: [],
      },
    ],
  },
];
