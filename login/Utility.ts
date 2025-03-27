import { IExecSyncResult } from "azure-pipelines-task-lib/toolrunner";
import tl = require('azure-pipelines-task-lib/task');

export class Utility {
    public static throwIfError(resultOfToolExecution: IExecSyncResult, errormsg?: string): void {
        if (resultOfToolExecution.code != 0) {
            tl.error("Error Code: [" + resultOfToolExecution.code + "]");
            if (errormsg) {
                tl.error("Error: " + errormsg);
            }
            throw resultOfToolExecution;
        }
    }
}