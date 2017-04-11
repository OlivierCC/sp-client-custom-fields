/**
 * @file PropertyFieldTermSetPicker.ts
 * Define a custom field of type PropertyFieldTermSetPicker for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IWebPartContext
} from '@microsoft/sp-webpart-base';
import PropertyFieldTermSetPickerHost, { IPropertyFieldTermSetPickerHostProps } from './PropertyFieldTermSetPickerHost';

/**
 * @interface
 * Generic Term Object (abstract interface)
 */
export interface ISPTermObject {
  Name: string;
  Guid: string;
  Identity: string;
  leaf: boolean;
  children?: ISPTermObject[];
  collapsed?: boolean;
  type: string;
}

/**
 * @interface
 * Defines a SharePoint Term Store
 */
export interface ISPTermStore extends ISPTermObject {
  IsOnline: boolean;
  WorkingLanguage: string;
  DefaultLanguage: string;
  Languages: string[];
}

/**
 * @interface
 * Defines an array of Term Stores
 */
export interface ISPTermStores extends Array<ISPTermStore> {
}

/**
 * @interface
 * Defines a Term Store Group of term sets
 */
export interface ISPTermGroup extends ISPTermObject {
  IsSiteCollectionGroup: boolean;
  IsSystemGroup: boolean;
  CreatedDate: string;
  LastModifiedDate: string;
}

/**
 * @interface
 * Array of Term Groups
 */
export interface ISPTermGroups extends Array<ISPTermGroup> {
}

/**
 * @interface
 * Defines a Term Set
 */
export interface ISPTermSet extends ISPTermObject {
  CustomSortOrder: string;
  IsAvailableForTagging: boolean;
  Owner: string;
  Contact: string;
  Description: string;
  IsOpenForTermCreation: boolean;
  TermStoreGuid: string;
}

/**
 * @interface
 * Array of Term Sets
 */
export interface ISPTermSets extends Array<ISPTermSet> {
}


/**
 * @interface
 * Public properties of the PropertyFieldTermSetPicker custom field
 *
 */
export interface IPropertyFieldTermSetPickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * TermSet Picker Panel title
   */
  panelTitle: string;
  /**
   * @var
   * Defines if the user can select only one or many term sets. Default value is false.
   *
   */
  allowMultipleSelections?: boolean;
  /**
   * @var
   * Defines the selected by default term sets.
   */
  initialValues?: ISPTermSets;
  /**
   * @var
   * Indicator to define if the system Groups are exclude. Default is false.
   */
  excludeSystemGroup?: boolean;
  /**
   * @var
   * Indicates if the offline term stores must be exclude. Default is false.
   */
  excludeOfflineTermStores?: boolean;
  /**
   * @var
   * Restrict term sets that are available for tagging. Default is false.
   */
  displayOnlyTermSetsAvailableForTagging?: boolean;
  /**
   * @var
   * WebPart's context
   */
  context: IWebPartContext;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected value changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
   * @var
   * Parent Web Part properties
   */
  properties: any;
  /**
   * @var
   * An UNIQUE key indicates the identity of this control
   */
  key?: string;
  /**
   * Whether the property pane field is enabled or not.
   */
  disabled?: boolean;
  /**
   * The method is used to get the validation error message and determine whether the input value is valid or not.
   *
   *   When it returns string:
   *   - If valid, it returns empty string.
   *   - If invalid, it returns the error message string and the text field will
   *     show a red border and show an error message below the text field.
   *
   *   When it returns Promise<string>:
   *   - The resolved value is display as error message.
   *   - The rejected, the value is thrown away.
   *
   */
   onGetErrorMessage?: (value: ISPTermSets) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldTermSetPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldTermSetPicker.
 *
 */
export interface IPropertyFieldTermSetPickerPropsInternal extends IPropertyFieldTermSetPickerProps {
  label: string;
  targetProperty: string;
  panelTitle: string;
  allowMultipleSelections?: boolean;
  initialValues?: ISPTermSets;
  excludeSystemGroup?: boolean;
  excludeOfflineTermStores?: boolean;
  displayOnlyTermSetsAvailableForTagging?: boolean;
  context: IWebPartContext;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  key: string;
  disabled?: boolean;
  onGetErrorMessage?: (value: ISPTermSets) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldTermSetPicker object
 *
 */
class PropertyFieldTermSetPickerBuilder implements IPropertyPaneField<IPropertyFieldTermSetPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldTermSetPickerPropsInternal;

  //Custom properties label: string;
  private label: string;
  private context: IWebPartContext;
  private allowMultipleSelections: boolean = false;
  private initialValues: ISPTermSets = [];
  private excludeSystemGroup: boolean = false;
  private excludeOfflineTermStores: boolean = false;
  private displayOnlyTermSetsAvailableForTagging: boolean = false;
  private panelTitle: string;

  public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {}
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: ISPTermSets) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldTermSetPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.label = _properties.label;
    this.context = _properties.context;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    if (_properties.allowMultipleSelections !== undefined)
      this.allowMultipleSelections = _properties.allowMultipleSelections;
    if (_properties.initialValues !== undefined)
      this.initialValues = _properties.initialValues;
    if (_properties.excludeSystemGroup !== undefined)
      this.excludeSystemGroup = _properties.excludeSystemGroup;
    if (_properties.excludeOfflineTermStores !== undefined)
      this.excludeOfflineTermStores = _properties.excludeOfflineTermStores;
    if (_properties.displayOnlyTermSetsAvailableForTagging !== undefined)
      this.displayOnlyTermSetsAvailableForTagging = _properties.displayOnlyTermSetsAvailableForTagging;
      this.panelTitle = _properties.panelTitle;
  }

  /**
   * @function
   * Renders the SPListPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldTermSetPickerHostProps> = React.createElement(PropertyFieldTermSetPickerHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      panelTitle: this.panelTitle,
      allowMultipleSelections: this.allowMultipleSelections,
      initialValues: this.initialValues,
      excludeSystemGroup: this.excludeSystemGroup,
      excludeOfflineTermStores: this.excludeOfflineTermStores,
      displayOnlyTermSetsAvailableForTagging: this.displayOnlyTermSetsAvailableForTagging,
      context: this.context,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
      disabled: this.disabled,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);
  }

  /**
   * @function
   * Disposes the current object
   */
  private dispose(elem: HTMLElement): void {

  }

}

/**
 * @function
 * Helper method to create a SPList Picker on the PropertyPane.
 * @param targetProperty - Target property the SharePoint list picker is associated to.
 * @param properties - Strongly typed SPList Picker properties.
 */
export function PropertyFieldTermSetPicker(targetProperty: string, properties: IPropertyFieldTermSetPickerProps): IPropertyPaneField<IPropertyFieldTermSetPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldTermSetPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      panelTitle: properties.panelTitle,
      allowMultipleSelections: properties.allowMultipleSelections,
      initialValues: properties.initialValues,
      excludeSystemGroup: properties.excludeSystemGroup,
      excludeOfflineTermStores: properties.excludeOfflineTermStores,
      displayOnlyTermSetsAvailableForTagging: properties.displayOnlyTermSetsAvailableForTagging,
      context: properties.context,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldTermSetPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldTermSetPickerBuilder(targetProperty, newProperties);
}
