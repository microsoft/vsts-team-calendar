import Contracts = require("TFS/Build/Contracts");
import VSS_WebApi = require("VSS/WebApi/RestClient");
export declare class BuildHttpClient extends VSS_WebApi.VssHttpClient {
    static serviceInstanceId: string;
    constructor(rootRequestPath: string);
    /**
     * Associates an artifact with a build
     *
     * @param {Contracts.BuildArtifact} artifact
     * @param {number} buildId
     * @param {string} project - Project ID or project name
     * @return IPromise<Contracts.BuildArtifact>
     */
    createArtifact(artifact: Contracts.BuildArtifact, buildId: number, project?: string): IPromise<Contracts.BuildArtifact>;
    /**
     * Gets a specific artifact for a build
     *
     * @param {number} buildId
     * @param {string} artifactName
     * @param {string} project - Project ID or project name
     * @return IPromise<Contracts.BuildArtifact>
     */
    getArtifact(buildId: number, artifactName: string, project?: string): IPromise<Contracts.BuildArtifact>;
    /**
     * Gets all artifacts for a build
     *
     * @param {number} buildId
     * @param {string} project - Project ID or project name
     * @return IPromise<Contracts.BuildArtifact[]>
     */
    getArtifacts(buildId: number, project?: string): IPromise<Contracts.BuildArtifact[]>;
    /**
     * @param {string} project
     * @param {number} definitionId
     * @param {string} branchName
     * @return IPromise<string>
     */
    getBadge(project: string, definitionId: number, branchName?: string): IPromise<string>;
    /**
     * Deletes a build
     *
     * @param {number} buildId
     * @param {string} project - Project ID or project name
     * @return IPromise<void>
     */
    deleteBuild(buildId: number, project?: string): IPromise<void>;
    /**
     * Gets a build
     *
     * @param {number} buildId
     * @param {string} project - Project ID or project name
     * @param {string} propertyFilters - A comma-delimited list of properties to include in the results
     * @return IPromise<Contracts.Build>
     */
    getBuild(buildId: number, project?: string, propertyFilters?: string): IPromise<Contracts.Build>;
    /**
     * Gets builds
     *
     * @param {string} project - Project ID or project name
     * @param {number[]} definitions - A comma-delimited list of definition ids
     * @param {number[]} queues - A comma-delimited list of queue ids
     * @param {string} buildNumber
     * @param {Date} minFinishTime
     * @param {Date} maxFinishTime
     * @param {string} requestedFor
     * @param {Contracts.BuildReason} reasonFilter
     * @param {Contracts.BuildStatus} statusFilter
     * @param {Contracts.BuildResult} resultFilter
     * @param {string[]} tagFilters - A comma-delimited list of tags
     * @param {string[]} properties - A comma-delimited list of properties to include in the results
     * @param {Contracts.DefinitionType} type - The definition type
     * @param {number} top - The maximum number of builds to retrieve
     * @param {string} continuationToken
     * @return IPromise<Contracts.Build[]>
     */
    getBuilds(project?: string, definitions?: number[], queues?: number[], buildNumber?: string, minFinishTime?: Date, maxFinishTime?: Date, requestedFor?: string, reasonFilter?: Contracts.BuildReason, statusFilter?: Contracts.BuildStatus, resultFilter?: Contracts.BuildResult, tagFilters?: string[], properties?: string[], type?: Contracts.DefinitionType, top?: number, continuationToken?: string): IPromise<Contracts.Build[]>;
    /**
     * Queues a build
     *
     * @param {Contracts.Build} build
     * @param {string} project - Project ID or project name
     * @param {boolean} ignoreWarnings
     * @return IPromise<Contracts.Build>
     */
    queueBuild(build: Contracts.Build, project?: string, ignoreWarnings?: boolean): IPromise<Contracts.Build>;
    /**
     * Updates a build
     *
     * @param {Contracts.Build} build
     * @param {number} buildId
     * @param {string} project - Project ID or project name
     * @return IPromise<Contracts.Build>
     */
    updateBuild(build: Contracts.Build, buildId: number, project?: string): IPromise<Contracts.Build>;
    /**
     * Gets the changes associated with a build
     *
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @param {number} top - The maximum number of changes to return
     * @return IPromise<Contracts.Change[]>
     */
    getBuildCommits(project: string, buildId: number, top?: number): IPromise<Contracts.Change[]>;
    /**
     * Gets a controller
     *
     * @param {number} controllerId
     * @return IPromise<Contracts.BuildController>
     */
    getBuildController(controllerId: number): IPromise<Contracts.BuildController>;
    /**
     * Gets controller, optionally filtered by name
     *
     * @param {string} name
     * @return IPromise<Contracts.BuildController[]>
     */
    getBuildControllers(name?: string): IPromise<Contracts.BuildController[]>;
    /**
     * Creates a new definition
     *
     * @param {Contracts.BuildDefinition} definition
     * @param {string} project - Project ID or project name
     * @param {number} definitionToCloneId
     * @param {number} definitionToCloneRevision
     * @return IPromise<Contracts.BuildDefinition>
     */
    createDefinition(definition: Contracts.BuildDefinition, project?: string, definitionToCloneId?: number, definitionToCloneRevision?: number): IPromise<Contracts.BuildDefinition>;
    /**
     * Deletes a definition
     *
     * @param {number} definitionId
     * @param {string} project - Project ID or project name
     * @return IPromise<void>
     */
    deleteDefinition(definitionId: number, project?: string): IPromise<void>;
    /**
     * Gets a definition, optionally at a specific revision
     *
     * @param {number} definitionId
     * @param {string} project - Project ID or project name
     * @param {number} revision
     * @param {string[]} propertyFilters
     * @return IPromise<Contracts.DefinitionReference>
     */
    getDefinition(definitionId: number, project?: string, revision?: number, propertyFilters?: string[]): IPromise<Contracts.DefinitionReference>;
    /**
     * Gets definitions, optionally filtered by name
     *
     * @param {string} project - Project ID or project name
     * @param {string} name
     * @param {Contracts.DefinitionType} type
     * @return IPromise<Contracts.DefinitionReference[]>
     */
    getDefinitions(project?: string, name?: string, type?: Contracts.DefinitionType): IPromise<Contracts.DefinitionReference[]>;
    /**
     * Updates an existing definition
     *
     * @param {Contracts.BuildDefinition} definition
     * @param {number} definitionId
     * @param {string} project - Project ID or project name
     * @return IPromise<Contracts.BuildDefinition>
     */
    updateDefinition(definition: Contracts.BuildDefinition, definitionId: number, project?: string): IPromise<Contracts.BuildDefinition>;
    /**
     * Gets logs for a build
     *
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @return IPromise<Contracts.BuildLog[]>
     */
    getBuildLogs(project: string, buildId: number): IPromise<Contracts.BuildLog[]>;
    /**
     * @return IPromise<Contracts.BuildOptionDefinition[]>
     */
    getBuildOptionDefinitions(): IPromise<Contracts.BuildOptionDefinition[]>;
    /**
     * Creates a build queue
     *
     * @param {Contracts.AgentPoolQueue} queue
     * @return IPromise<Contracts.AgentPoolQueue>
     */
    createQueue(queue: Contracts.AgentPoolQueue): IPromise<Contracts.AgentPoolQueue>;
    /**
     * Deletes a build queue
     *
     * @param {number} id
     * @return IPromise<void>
     */
    deleteQueue(id: number): IPromise<void>;
    /**
     * Gets a queue
     *
     * @param {number} controllerId
     * @return IPromise<Contracts.AgentPoolQueue>
     */
    getAgentPoolQueue(controllerId: number): IPromise<Contracts.AgentPoolQueue>;
    /**
     * Gets queues, optionally filtered by name
     *
     * @param {string} name
     * @return IPromise<Contracts.AgentPoolQueue[]>
     */
    getQueues(name?: string): IPromise<Contracts.AgentPoolQueue[]>;
    /**
     * Gets revisions of a definition
     *
     * @param {string} project - Project ID or project name
     * @param {number} definitionId
     * @return IPromise<Contracts.BuildDefinitionRevision[]>
     */
    getDefinitionRevisions(project: string, definitionId: number): IPromise<Contracts.BuildDefinitionRevision[]>;
    /**
     * @return IPromise<Contracts.BuildSettings>
     */
    getBuildSettings(): IPromise<Contracts.BuildSettings>;
    /**
     * Updates the build settings
     *
     * @param {Contracts.BuildSettings} settings
     * @return IPromise<Contracts.BuildSettings>
     */
    updateBuildSettings(settings: Contracts.BuildSettings): IPromise<Contracts.BuildSettings>;
    /**
     * Adds a tag to a build
     *
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @param {string} tag
     * @return IPromise<string[]>
     */
    addBuildTag(project: string, buildId: number, tag: string): IPromise<string[]>;
    /**
     * Adds tag to a build
     *
     * @param {string[]} tags
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @return IPromise<string[]>
     */
    addBuildTags(tags: string[], project: string, buildId: number): IPromise<string[]>;
    /**
     * Deletes a tag from a build
     *
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @param {string} tag
     * @return IPromise<string[]>
     */
    deleteBuildTag(project: string, buildId: number, tag: string): IPromise<string[]>;
    /**
     * Gets the tags for a build
     *
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @return IPromise<string[]>
     */
    getBuildTags(project: string, buildId: number): IPromise<string[]>;
    /**
     * @param {string} project - Project ID or project name
     * @return IPromise<string[]>
     */
    getTags(project: string): IPromise<string[]>;
    /**
     * Deletes a definition template
     *
     * @param {string} project - Project ID or project name
     * @param {string} templateId
     * @return IPromise<void>
     */
    deleteTemplate(project: string, templateId: string): IPromise<void>;
    /**
     * Gets definition template filtered by id
     *
     * @param {string} project - Project ID or project name
     * @param {string} templateId
     * @return IPromise<Contracts.BuildDefinitionTemplate>
     */
    getTemplate(project: string, templateId: string): IPromise<Contracts.BuildDefinitionTemplate>;
    /**
     * @param {string} project - Project ID or project name
     * @return IPromise<Contracts.BuildDefinitionTemplate[]>
     */
    getTemplates(project: string): IPromise<Contracts.BuildDefinitionTemplate[]>;
    /**
     * Saves a definition template
     *
     * @param {Contracts.BuildDefinitionTemplate} template
     * @param {string} project - Project ID or project name
     * @param {string} templateId
     * @return IPromise<Contracts.BuildDefinitionTemplate>
     */
    saveTemplate(template: Contracts.BuildDefinitionTemplate, project: string, templateId: string): IPromise<Contracts.BuildDefinitionTemplate>;
    /**
     * Gets details for a build
     *
     * @param {string} project - Project ID or project name
     * @param {number} buildId
     * @param {string} timelineId
     * @return IPromise<Contracts.Timeline>
     */
    getBuildTimeline(project: string, buildId: number, timelineId?: string): IPromise<Contracts.Timeline>;
}
