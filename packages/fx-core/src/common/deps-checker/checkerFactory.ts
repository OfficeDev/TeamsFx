// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DotnetChecker } from "./internal/dotnetChecker";
import { DepsChecker } from "./depsChecker";
import { ProjectNodeChecker, LtsNodeChecker } from "./internal/nodeChecker";
import { FuncToolChecker } from "./internal/funcToolChecker";
import { DepsType } from "./depsChecker";
import { VxTestAppChecker } from "./internal/vxTestAppChecker";

export class CheckerFactory {
  public static createChecker(type: DepsType): DepsChecker {
    switch (type) {
      case DepsType.ProjectNode:
        return new ProjectNodeChecker();
      case DepsType.LtsNode:
        return new LtsNodeChecker();
      case DepsType.Dotnet:
        return new DotnetChecker();
      case DepsType.FuncCoreTools:
        return new FuncToolChecker();
      case DepsType.VxTestApp:
        return new VxTestAppChecker();
      default:
        throw Error("dependency type is undefined");
    }
  }
}
