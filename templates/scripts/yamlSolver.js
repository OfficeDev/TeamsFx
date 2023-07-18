const fs = require("fs-extra");
const path = require("path");
const utils = require("./utils");

// The solver is called by the following command:
// > node yamlSolver.js <action> <constraintFilePath>

// The solver support two actions
// 1. apply the constraints to solution
// 2. verify the constraints are satisfied.
const Action = {
  APPLY: "apply",
  VERIFY: "verify",
};

// The constraints are defined in mustache files.
// default constraints folders
const mustacheFolder = path.resolve(__dirname, "..", "constraints", "yml", "templates");
const solutionFolder = path.resolve(__dirname, "..");

// example:  " key1: value, key2 " => { key1: value, key2: true }
function strToObj(str) {
  const properties = str.split(",");
  let obj = {};
  properties.forEach(function (property) {
    if (property.includes(":")) {
      const tup = property.split(":");
      obj[tup[0].trim()] = tup[1].trim();
    } else {
      obj[property.trim()] = true;
    }
  });
  return obj;
}

// read all yml files and mustache files in folder as mustache variables
function generateVariablesFromSnippets(dir) {
  let result = {};
  utils.filterYmlFiles(dir).map((file) => {
    const yml = fs.readFileSync(file, "utf8");
    result = { ...result, ...{ [path.basename(file, ".yml")]: yml } };
  });
  utils.filterMustacheFiles(dir).map((file) => {
    const mustache = fs.readFileSync(file, "utf8");
    result = {
      ...result,
      ...{
        [path.basename(file, ".mustache")]: function () {
          return function (text) {
            return utils.renderMustache(mustache, strToObj(text)).trimEnd();
          };
        },
      },
    };
  });
  return result;
}

function* solveMustache(mustachePaths) {
  for (const mustachePath of mustachePaths) {
    const template = fs.readFileSync(mustachePath, "utf8");
    const variables = generateVariablesFromSnippets(
      path.resolve(__dirname, "..", "constraints", "yml", "snippets")
    );
    const solution = utils.renderMustache(template, variables);
    yield { mustachePath, solution };
  }
}

function validateMustachePath(mustachePath) {
  // no input, return all mustache files
  if (!mustachePath) {
    return utils.filterMustacheFiles(mustacheFolder);
  }
  // input is a folder, return all mustache files in folder
  if (fs.lstatSync(mustachePath).isDirectory()) {
    return utils.filterMustacheFiles(mustachePath);
  }
  if (!mustachePath.endsWith(".mustache")) {
    throw new Error("Invalid mustache file path");
  }
  if (!fs.existsSync(mustachePath)) {
    throw new Error("Invalid path");
  }
  // return input mustache file
  return [mustachePath];
}

function validateAction(action) {
  if (!Object.values(Action).includes(action)) {
    throw new Error(`Invalid action. Must be either ${Object.values(Action)}`);
  }
  return action;
}

function parseInput() {
  return {
    action: validateAction(process.argv[2]),
    mustachePaths: validateMustachePath(process.argv[3]),
  };
}

function main({ action, mustachePaths }) {
  let SAT = Action.VERIFY === action;
  for (const { mustachePath, solution } of solveMustache(mustachePaths)) {
    const solutionPath = path.resolve(
      solutionFolder,
      path.dirname(path.relative(mustacheFolder, mustachePath)),
      path.basename(mustachePath, ".mustache") + ".yml.tpl"
    );
    switch (action) {
      case Action.APPLY:
        fs.writeFileSync(solutionPath, solution);
        break;
      case Action.VERIFY:
        const expected = fs.readFileSync(solutionPath, "utf8");
        const assertion = solution.replaceAll("\r\n", "\n") === expected.replaceAll("\r\n", "\n");
        console.assert(
          assertion,
          `${solutionPath} is not satisfied with the constraint ${mustachePath}`
        );
        SAT = SAT && assertion;
        break;
    }
  }

  if (SAT) {
    console.log("All constraints are satisfied");
  }
}

main(parseInput());
