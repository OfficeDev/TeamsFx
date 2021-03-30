import { Result } from 'neverthrow';
import { FxError } from '../error';
export interface TreeItem {
    commandId: string;
    label: string;
    callback?: (args: any) => Promise<Result<null, FxError>>;
    parent?: TreeCategory | string;
    contextValue?: string;
    icon?: string;
    subTreeItems?: TreeItem[];
}
export interface TreeProvider {
    refresh: (tree: TreeItem[]) => Promise<Result<null, FxError>>;
    add: (tree: TreeItem[]) => Promise<Result<null, FxError>>;
    remove: (tree: TreeItem[]) => Promise<Result<null, FxError>>;
}
export declare enum TreeCategory {
    GettingStarted = 0,
    Account = 1,
    Feedback = 2,
    Project = 3,
    Provison = 4
}
//# sourceMappingURL=tree.d.ts.map