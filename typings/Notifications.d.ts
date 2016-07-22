// Type definitions for Microsoft Visual Studio Services v104.20160722.1311
// Project: https://www.visualstudio.com/integrate/extensions/overview
// Definitions by: Microsoft <vsointegration@microsoft.com>

/// <reference path='vss.d.ts' />
declare module "Notifications/Contracts" {
import VSS_Common_Contracts = require("VSS/WebApi/Contracts");
import VSS_FormInput_Contracts = require("VSS/Common/Contracts/FormInput");
export interface ActorFilter extends RoleBasedFilter {
}
export interface ArtifactFilter extends BaseSubscriptionFilter {
    artifactId: string;
    artifactType: string;
    artifactUri: string;
    type: string;
}
export interface ArtifactSubscription {
    artifactId: string;
    artifactType: string;
    subscriptionId: number;
}
export interface BaseSubscriptionFilter {
    eventType: string;
    type: string;
}
export interface ChatRoomSubscriptionChannel extends SubscriptionChannelWithAddress {
    type: string;
}
export interface EmailHtmlSubscriptionChannel extends SubscriptionChannelWithAddress {
    type: string;
}
export interface EmailPlaintextSubscriptionChannel extends SubscriptionChannelWithAddress {
    type: string;
}
export interface ExpressionFilter extends BaseSubscriptionFilter {
    criteria: ExpressionFilterModel;
    type: string;
}
/**
 * Subscription Filter Clause represents a single clause in a subscription filter e.g. If the subscription has the following criteria "Project Name = [Current Project] AND Assigned To = [Me] it will be represented as two Filter Clauses Clause 1: Index = 1, Logical Operator: NULL  , FieldName = 'Project Name', Operator = '=', Value = '[Current Project]' Clause 2: Index = 2, Logical Operator: 'AND' , FieldName = 'Assigned To' , Operator = '=', Value = '[Me]'
 */
export interface ExpressionFilterClause {
    fieldName: string;
    /**
     * The order in which this clause appeared in the filter query
     */
    index: number;
    /**
     * Logical Operator 'AND', 'OR' or NULL (only for the first clause in the filter)
     */
    logicalOperator: string;
    operator: string;
    value: string;
}
/**
 * Represents a hierarchy of SubscritionFilterClauses that have been grouped together through either adding a group in the WebUI or using parethesis in the Subscription condition string
 */
export interface ExpressionFilterGroup {
    /**
     * The index of the last FilterClause in this group
     */
    end: number;
    /**
     * Level of the group, since groups can be nested for each nested group the level will increase by 1
     */
    level: number;
    /**
     * The index of the first FilterClause in this group
     */
    start: number;
}
export interface ExpressionFilterModel {
    /**
     * Flat list of clauses in this subscription
     */
    clauses: ExpressionFilterClause[];
    /**
     * Grouping of clauses in the subscription
     */
    groups: ExpressionFilterGroup[];
    /**
     * Max depth of the Subscription tree
     */
    maxGroupLevel: number;
}
export interface FieldInputValues extends VSS_FormInput_Contracts.InputValues {
    operators: number[];
}
export interface FieldValuesQuery extends VSS_FormInput_Contracts.InputValuesQuery {
    inputValues: FieldInputValues[];
    scope: string;
}
export interface IgnoreFilter extends RoleBasedFilter {
}
export interface ISubscriptionChannel {
    type: string;
}
export interface ISubscriptionFilter {
    eventType: string;
    type: string;
}
export interface MessageQueueSubscriptionChannel {
    type: string;
}
export interface NotificationDataProviderData {
    isTeam: boolean;
    mapCategoryNameToSubscriptionTemplates: {
        [key: string]: NotificationSubscriptionTemplate[];
    };
    mapEventTypeToCategoryName: {
        [key: string]: string;
    };
    mapEventTypeToEventAlias: {
        [key: string]: string;
    };
    mapEventTypeToEventInfo: {
        [key: string]: NotificationEventTypeInformation;
    };
    mapEventTypeToPublisherId: {
        [key: string]: string;
    };
    mapProjectIdToProjectName: {
        [key: string]: string;
    };
    publishers: {
        [key: string]: NotificationEventPublisher;
    };
    subscriptions: NotificationSubscription[];
}
/**
 * Encapsulates the properties of a filterable field. A filterable field is a field in an event that can used to filter notifications for a certain event type.
 */
export interface NotificationEventField {
    /**
     * Gets or sets the type of this field.
     */
    fieldType: NotificationEventFieldType;
    /**
     * Gets or sets the unique identifier of this field.
     */
    id: string;
    /**
     * Gets or sets if this field can be used as a role or an actor. This property only applies to identity fields
     */
    isRole: boolean;
    /**
     * Gets or sets the name of this field.
     */
    name: string;
    /**
     * Gets or sets the path to the field in the event object. This path can be either Json Path or XPath, depending on if the event will be serialized into Json or XML
     */
    path: string;
}
/**
 * Encapsulates the properties of a field type. It describes the data type of a field, the operators it support and how to populate it in the UI
 */
export interface NotificationEventFieldType {
    /**
     * Gets or sets the unique identifier of this field type.
     */
    id: string;
    operatorConstraints: OperatorConstraint[];
    /**
     * Gets or sets the list of operators that this type supports.
     */
    operators: string[];
    /**
     * Gets or sets the value definition of this field like the getValuesMethod and template to display in the UI
     */
    value: ValueDefinition;
}
/**
 * Encapsulates the properties of a notification event publisher.
 */
export interface NotificationEventPublisher {
    id: string;
    subscriptionManagementInfo: SubscriptionManagement;
    url: string;
}
/**
 * Encapsulates the properties of a category. A category will be used by the UI to group event types
 */
export interface NotificationEventTypeCategory {
    /**
     * Gets or sets the unique identifier of this category.
     */
    id: string;
    /**
     * Gets or sets the friendly name of this category.
     */
    name: string;
}
/**
 * Encapsulates the properties of an event type. It defines the fields, that can be used for filtering, for that event type.
 */
export interface NotificationEventTypeInformation {
    category: NotificationEventTypeCategory;
    eventPublisher: NotificationEventPublisher;
    fields: {
        [key: string]: NotificationEventField;
    };
    /**
     * Gets or sets the unique identifier of this event definition.
     */
    id: string;
    /**
     * Gets or sets the name of this event definition.
     */
    name: string;
    /**
     * Gets or sets the rest end point to get this event type details (fields, fields types)
     */
    url: string;
}
export interface NotificationSubscription extends NotificationSubscriptionBase {
    _links: any;
    adminConfig: SubscriptionAdminConfig;
    channel: ISubscriptionChannel;
    flags: SubscriptionFlags;
    lastModifiedBy: VSS_Common_Contracts.IdentityRef;
    modifiedDate: Date;
    scope: string;
    status: SubscriptionStatus;
    statusMessage: string;
    subscriber: VSS_Common_Contracts.IdentityRef;
    subscriptionProvider: string;
    url: string;
    userConfig: SubscriptionUserConfig;
}
export interface NotificationSubscriptionBase {
    description: string;
    filter: ISubscriptionFilter;
    id: string;
}
export interface NotificationSubscriptionTemplate extends NotificationSubscriptionBase {
    notificationEventInformation: NotificationEventTypeInformation;
    type: SubscriptionTemplateType;
}
/**
 * Encapsulates the properties of an operator constraint. An operator constraint defines if some operator is available only for specific scope like a project scope.
 */
export interface OperatorConstraint {
    operator: string;
    /**
     * Gets or sets the list of operators that this type supports.
     */
    scope: string[];
}
export interface RoleBasedFilter extends ExpressionFilter {
    exclusions: string[];
    inclusions: string[];
}
export interface ServiceHooksSubscriptionChannel {
    type: string;
}
export interface SoapSubscriptionChannel extends SubscriptionChannelWithAddress {
    type: string;
}
export enum SubscriptionActions {
    None = 0,
    Edit = 1,
    Delete = 2,
    UserOptOut = 4,
    EditAdminConfig = 8,
}
export interface SubscriptionAdminConfig {
    blockUserDisable: boolean;
    enabled: boolean;
}
export interface SubscriptionChannelWithAddress {
    address: string;
    type: string;
}
export enum SubscriptionFlags {
    None = 0,
    TeamSubscription = 1,
    ContributedSubscription = 2,
}
/**
 * Encapsulates the properties needed to manage subscriptions, opt in and out of subscriptions.
 */
export interface SubscriptionManagement {
    serviceInstanceType: string;
    url: string;
}
/**
 * For quering, back in response
 */
export interface SubscriptionQuery {
    conditions: SubscriptionQueryCondition[];
}
export interface SubscriptionQueryCondition {
    filter: ISubscriptionFilter;
    scope: string;
    subscriber: string;
    subscriptionType: SubscriptionType;
}
export enum SubscriptionStatus {
    DisabledMissingIdentity = -6,
    DusabledInvalidRoleExpression = -5,
    DisabledInvalidPathClause = -4,
    DisabledAsDuplicateOfDefault = -3,
    DisabledByAdmin = -2,
    DisabledByUser = -1,
    Enabled = 0,
    EnabledOnProbation = 1,
}
export enum SubscriptionTemplateType {
    User = 0,
    Team = 1,
    Both = 2,
    None = 3,
}
export enum SubscriptionType {
    Default = 0,
    Shared = 1,
}
/**
 * Default subscriptions for now. return status for user.
 */
export interface SubscriptionUserConfig {
    actions: SubscriptionActions;
    enabled: boolean;
}
export interface UnsupportedFilter extends BaseSubscriptionFilter {
    type: string;
}
export interface UnsupportedSubscriptionChannel {
    type: string;
}
export interface UserSubscriptionChannel extends SubscriptionChannelWithAddress {
    type: string;
}
/**
 * Encapsulates the properties of a field value definition. It has the information needed to retrieve the list of possible values for a certain field and how to handle that field values in the UI. This information includes what type of object this value represents, which property to use for UI display and which property to use for saving the subscription
 */
export interface ValueDefinition {
    /**
     * Gets or sets the data source.
     */
    dataSource: any[];
    /**
     * Gets or sets the rest end point.
     */
    endPoint: string;
    /**
     * Gets or sets the result template.
     */
    resultTemplate: string;
}
export var TypeInfo: {
    NotificationDataProviderData: any;
    NotificationSubscription: any;
    NotificationSubscriptionTemplate: any;
    SubscriptionActions: {
        enumValues: {
            "none": number;
            "edit": number;
            "delete": number;
            "userOptOut": number;
            "editAdminConfig": number;
        };
    };
    SubscriptionFlags: {
        enumValues: {
            "none": number;
            "teamSubscription": number;
            "contributedSubscription": number;
        };
    };
    SubscriptionQuery: any;
    SubscriptionQueryCondition: any;
    SubscriptionStatus: {
        enumValues: {
            "disabledMissingIdentity": number;
            "dusabledInvalidRoleExpression": number;
            "disabledInvalidPathClause": number;
            "disabledAsDuplicateOfDefault": number;
            "disabledByAdmin": number;
            "disabledByUser": number;
            "enabled": number;
            "enabledOnProbation": number;
        };
    };
    SubscriptionTemplateType: {
        enumValues: {
            "user": number;
            "team": number;
            "both": number;
            "none": number;
        };
    };
    SubscriptionType: {
        enumValues: {
            "default": number;
            "shared": number;
        };
    };
    SubscriptionUserConfig: any;
};
}
declare module "Notifications/notifications.test" {
}
declare module "Notifications/Resources" {
export module FollowsCustomerIntelligenceConstants {
    var Feature: string;
    var FollowAction: string;
    var UnfollowAction: string;
}
export module FollowsResource {
    var LocationId: string;
    var Resource: string;
}
export module NotificationResourceIds {
    var AreaName: string;
}
export module NotificationsCustomerIntelligenceConstants {
    var Area: string;
}
}
declare module "Notifications/RestClient" {
import Contracts = require("Notifications/Contracts");
import VSS_Common_Contracts = require("VSS/WebApi/Contracts");
import VSS_WebApi = require("VSS/WebApi/RestClient");
export class CommonMethods2_2To3 extends VSS_WebApi.VssHttpClient {
    static serviceInstanceId: string;
    protected followsApiVersion: string;
    constructor(rootRequestPath: string, options?: VSS_WebApi.IVssHttpClientOptions);
    /**
     * [Preview API]
     *
     * @param {number} subscriptionId
     * @return IPromise<void>
     */
    unfollowArtifact(subscriptionId: number): IPromise<void>;
    /**
     * [Preview API]
     *
     * @return IPromise<Contracts.ArtifactSubscription[]>
     */
    getArtifactSubscriptions(): IPromise<Contracts.ArtifactSubscription[]>;
    /**
     * [Preview API]
     *
     * @param {string} artifactType
     * @param {string} artifactId
     * @return IPromise<Contracts.ArtifactSubscription>
     */
    getArtifactSubscription(artifactType: string, artifactId: string): IPromise<Contracts.ArtifactSubscription>;
    /**
     * [Preview API]
     *
     * @param {Contracts.ArtifactSubscription} artifact
     * @return IPromise<Contracts.ArtifactSubscription>
     */
    followArtifact(artifact: Contracts.ArtifactSubscription): IPromise<Contracts.ArtifactSubscription>;
}
/**
 * @exemptedapi
 */
export class NotificationHttpClient3 extends CommonMethods2_2To3 {
    constructor(rootRequestPath: string, options?: VSS_WebApi.IVssHttpClientOptions);
    /**
     * [Preview API]
     *
     * @param {VSS_Common_Contracts.VssNotificationEvent} notificationEvent
     * @return IPromise<VSS_Common_Contracts.VssNotificationEvent>
     */
    publishEvent(notificationEvent: VSS_Common_Contracts.VssNotificationEvent): IPromise<VSS_Common_Contracts.VssNotificationEvent>;
    /**
     * [Preview API]
     *
     * @param {Contracts.FieldValuesQuery} inputValuesQuery
     * @param {string} eventType
     * @return IPromise<Contracts.NotificationEventField[]>
     */
    queryEventTypes(inputValuesQuery: Contracts.FieldValuesQuery, eventType: string): IPromise<Contracts.NotificationEventField[]>;
    /**
     * [Preview API]
     *
     * @param {string} eventType
     * @return IPromise<Contracts.NotificationEventTypeInformation>
     */
    getEventType(eventType: string): IPromise<Contracts.NotificationEventTypeInformation>;
    /**
     * [Preview API]
     *
     * @param {string} publisherId
     * @return IPromise<Contracts.NotificationEventTypeInformation[]>
     */
    getEventTypes(publisherId?: string): IPromise<Contracts.NotificationEventTypeInformation[]>;
    /**
     * [Preview API]
     *
     * @param {Contracts.SubscriptionQuery} subscriptionQuery
     * @return IPromise<Contracts.NotificationSubscription[]>
     */
    subscriptionQuery(subscriptionQuery: Contracts.SubscriptionQuery): IPromise<Contracts.NotificationSubscription[]>;
    /**
     * [Preview API]
     *
     * @param {Contracts.NotificationSubscription} notificationSubscription
     * @return IPromise<Contracts.NotificationSubscription>
     */
    createSubscription(notificationSubscription: Contracts.NotificationSubscription): IPromise<Contracts.NotificationSubscription>;
    /**
     * [Preview API]
     *
     * @param {string} subscriptionId
     * @return IPromise<void>
     */
    deleteSubscription(subscriptionId: string): IPromise<void>;
    /**
     * [Preview API]
     *
     * @param {string} subscriptionId
     * @return IPromise<Contracts.NotificationSubscription>
     */
    getSubscription(subscriptionId: string): IPromise<Contracts.NotificationSubscription>;
    /**
     * [Preview API]
     *
     * @param {Contracts.NotificationSubscription} subscriptionPatch
     * @param {string} subscriptionId
     * @return IPromise<Contracts.NotificationSubscription>
     */
    updateSubscription(subscriptionPatch: Contracts.NotificationSubscription, subscriptionId: string): IPromise<Contracts.NotificationSubscription>;
    /**
     * [Preview API]
     *
     * @return IPromise<Contracts.NotificationSubscriptionTemplate[]>
     */
    getSubscriptionTemplates(): IPromise<Contracts.NotificationSubscriptionTemplate[]>;
}
/**
 * @exemptedapi
 */
export class NotificationHttpClient2_2 extends CommonMethods2_2To3 {
    constructor(rootRequestPath: string, options?: VSS_WebApi.IVssHttpClientOptions);
}
export class NotificationHttpClient extends NotificationHttpClient3 {
    constructor(rootRequestPath: string, options?: VSS_WebApi.IVssHttpClientOptions);
}
/**
 * Gets an http client targeting the latest released version of the APIs.
 *
 * @return NotificationHttpClient2_2
 */
export function getClient(options?: VSS_WebApi.IVssHttpClientOptions): NotificationHttpClient2_2;
}
declare module "Notifications/Services" {
import Contracts = require("Notifications/Contracts");
import Service = require("VSS/Service");
export interface IFollowsTelemetryData {
    /**
    * Specifies the layer at which the action is performed
    */
    layer?: string;
    /**
    * Specific an area override
    */
    area?: string;
}
/**
 * Interface defining the arguments for the FOLLOWS_STATE_UPDATED event
 */
export interface IFollowsEventArgs {
    artifact: Contracts.ArtifactSubscription;
    isFollowing: boolean;
}
/**
 * Service to manage follows states
 */
export class FollowsService extends Service.VssService {
    private _httpClient;
    private _followsCache;
    static FOLLOWS_STATE_CHANGED: string;
    static FOLLOWS_STATE_CHANGING: string;
    private static LAYER;
    clearCache(): void;
    getSubscription(artifact: Contracts.ArtifactSubscription): IPromise<Contracts.ArtifactSubscription>;
    followArtifact(artifact: Contracts.ArtifactSubscription, telemetryData?: IFollowsTelemetryData): IPromise<Contracts.ArtifactSubscription>;
    unfollowArtifact(artifact: Contracts.ArtifactSubscription, telemetryData?: IFollowsTelemetryData): IPromise<Contracts.ArtifactSubscription>;
    refresh(artifact: Contracts.ArtifactSubscription): void;
    private _handlePromise(artifact, isFollow, promise);
    private _ensureInitialized();
    private _makePromiseKey(artifact);
    private _fireChanging(artifact, state);
    private _fireChanged(artifact, state);
    private _fireEvent(event, artifact, state);
    private _publishTelemetry(properties, action);
}
}
