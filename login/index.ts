import tl = require('azure-pipelines-task-lib/task');
import { Utility } from "./Utility";
import { getSystemAccessToken } from 'azure-pipelines-tasks-artifacts-common/webapi';
import { getHandlerFromToken, WebApi } from "azure-devops-node-api";
import { ITaskApi } from "azure-devops-node-api/TaskApi";
import { IExecSyncResult } from 'azure-pipelines-task-lib/toolrunner';

async function run() {
    try {
        const serviceConnectionInput: string | undefined = tl.getInput('serviceConnection', true);
        
        if (serviceConnectionInput == 'bad') {
            tl.setResult(tl.TaskResult.Failed, 'Bad input was given');
            return;
        }

        if (!serviceConnectionInput) {
            tl.setResult(tl.TaskResult.Failed, 'No service connection was configured on the task.');
            return;
        }

        login(serviceConnectionInput);
    }
    catch (err:any) {
        tl.setResult(tl.TaskResult.Failed, err.message);
    }
}

async function login(serviceConnection: string) {
    const authScheme = tl.getEndpointAuthorizationScheme(serviceConnection, true);

    if (!authScheme) {
        throw new Error("The authentication scheme of the service connection could not be determined.");
    }

    if (authScheme.toLowerCase() != "workloadidentityfederation") {
        throw new Error("The service connection must be of type 'Workload Identity Federation'.");
    }

    const servicePrincipalId = tl.getEndpointAuthorizationParameter(serviceConnection, "serviceprincipalid", false);
    const tenantId = tl.getEndpointAuthorizationParameter(serviceConnection, "tenantid", false);
    
    if (!servicePrincipalId) {
        throw new Error("The service principal Id of the service connection could not be determined.");
    }
    
    if (!tenantId) {
        throw new Error("The tenant Id of the service connection could not be determined.");
    }
    
    const federatedToken = await getIdToken(serviceConnection);

    const cliVersionResult: IExecSyncResult = tl.execSync("m365", "--version");

    if (cliVersionResult.code !== 0) {
        tl.execSync("npm", "install @pnp/cli-microsoft365@latest -g");
    }

    // check version 10.5.0

    let args = `login --authType federatedIdentity --appId ${servicePrincipalId} --tenant ${tenantId} --federated-token "${federatedToken}"`;

    Utility.throwIfError(tl.execSync("m365", args), "Login failed...");
}

async function getIdToken(serviceConnection: string): Promise<string | undefined> {
    // since node19 default node's GlobalAgent has timeout 5sec
    // keepAlive is set to true to avoid creating default node's GlobalAgent
    const webApiOptions = {
        keepAlive: true
    }
    const jobId = tl.getVariable("System.JobId");
    const planId = tl.getVariable("System.PlanId");
    const projectId = tl.getVariable("System.TeamProjectId");
    const hub = tl.getVariable("System.HostType");
    const uri = tl.getVariable("System.CollectionUri");
    const token = getSystemAccessToken();

    if (!uri) {
        throw new Error("The system variable 'System.CollectionUri' is not defined.");
    }

    if (!projectId) {
        throw new Error("The system variable 'System.TeamProjectId' is not defined.");
    }

    if (!hub) {
        throw new Error("The system variable 'System.HostType' is not defined.");
    }

    if (!planId) {
        throw new Error("The system variable 'System.PlanId' is not defined.");
    }

    if (!jobId) {
        throw new Error("The system variable 'System.JobId' is not defined.");
    }

    const authHandler = getHandlerFromToken(token);
    const connection = new WebApi(uri, authHandler, webApiOptions);
    const api: ITaskApi = await connection.getTaskApi();
    const response = await api.createOidcToken({}, projectId, hub, planId, jobId, serviceConnection);
    
    return response?.oidcToken;
}

run();