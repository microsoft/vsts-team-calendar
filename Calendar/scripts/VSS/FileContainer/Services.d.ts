/// <reference path="../References/VSS-Common.d.ts" />
import Service = require("VSS/Service");
export interface FileContainerPathInfo {
    containerId: number;
    path: string;
}
/**
* Service to manage file container data
*/
export declare class FileContainerService extends Service.VssService {
    private _httpClient;
    /**
     * Returns a list of file container items
     *
     * @param containerId The id of the container
     * @param scope The scope of the items
     * @param itemPath The path of the item within the container
     */
    beginGetItems(containerId: number, scope: string, itemPath: string): IPromise<any>;
    /**
     * Returns the file container info
     *
     * @param fileContainerPath The path of the container. For example, "#/12/drop".
     */
    parseContainerPath(fileContainerPath: string): FileContainerPathInfo;
}
