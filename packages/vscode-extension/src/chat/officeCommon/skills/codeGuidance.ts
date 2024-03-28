// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export function getCodeGenerateGuidance(host: string) {
  return `
  # Coding rules:
    - Code must be TypeScript compabible with ES2015.
    - Include type declarations in variable declaration, function return declaration, function argument declaration.
    - Use async/await over .then for Promise.
    - Use try-catch over .catch for Promise.
    - Use "fetch" over "XMLHttpRequest".
    - Don't use enum const. Like "Sunny", "Rainy", "Cloudy", or 0, 1, 2. Use enum instead.
    - Don't add "import" statement or "require" statement.
    - If The code use hypothetical service endpoint, must explain the response data structure with comment.
    - For multiple data types, using "as" keyword convert to single type.
    - Wrapped access to Office JavaScript object into the callback function of ${host}.run.
    - Run "$AccessObject".load("$PropertyName") before access the $Propery of the object.
    - Run "context.sync()" right after the $AccessObject.load() to sync the data.
  `;
}
