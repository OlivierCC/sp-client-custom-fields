/**
 * @file PropertyFieldSPListMultiplePicker.ts
 * Define a custom field of type PropertyFieldSPListMultiplePicker for
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
  PropertyPaneFieldType
} from '@microsoft/sp-webpart-base';
import PropertyFieldSPListMultiplePickerHost, { IPropertyFieldSPListMultiplePickerHostProps } from './PropertyFieldSPListMultiplePickerHost';
import { IWebPartContext} from '@microsoft/sp-webpart-base';

/**
 * @enum
 * Enumerated the sort order of the lists
 *
 */
export enum PropertyFieldSPListMultiplePickerOrderBy {
  Id = 0,
  Title = 1
}

/**
 * @interface
 * Public properties of the PropertyFieldSPListMultiplePicker custom field
 *
 */
export interface IPropertyFieldSPListMultiplePickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Parent web part context
   */
  context: IWebPartContext;
  /**
   * @var
   * Default selected values of the picker (must be a collection of list Ids)
   */
  selectedLists?: string[];
  /**
   * @var
   * Defines the base template number to filter the list kind
   */
  baseTemplate?: number;
  /**
   * @var
   * Defines if the hidden list are included or not
   */
  includeHidden?: boolean;
  /**
   * @var
   * Defines the lists order
   */
  orderBy?: PropertyFieldSPListMultiplePickerOrderBy;
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
   onGetErrorMessage?: (value: string[]) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldSPListMultiplePicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSPListMultiplePicker.
 *
 */
export interface IPropertyFieldSPListMultiplePickerPropsInternal extends IPropertyFieldSPListMultiplePickerProps {
  label: string;
  targetProperty: string;
  context: IWebPartContext;
  selectedLists?: string[];
  baseTemplate?: number;
  orderBy?: PropertyFieldSPListMultiplePickerOrderBy;
  includeHidden?: boolean;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  key: string;
  disabled?: boolean;
  onGetErrorMessage?: (value: string[]) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldSPListMultiplePicker object
 *
 */
class PropertyFieldSPListMultiplePickerBuilder implements IPropertyPaneField<IPropertyFieldSPListMultiplePickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSPListMultiplePickerPropsInternal;

  //Custom properties label: string;
  private label: string;
  private context: IWebPartContext;
  private selectedLists: string[];
  private baseTemplate: number;
  private orderBy: PropertyFieldSPListMultiplePickerOrderBy;
  private includeHidden: boolean;

  public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {}
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: string[]) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSPListMultiplePickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.label = _properties.label;
    this.context = _properties.context;
    this.selectedLists = _properties.selectedLists;
    this.baseTemplate = _properties.baseTemplate;
    this.orderBy = _properties.orderBy;
    this.includeHidden = _properties.includeHidden;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
  }

  /**
   * @function
   * Renders the SPListPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldSPListMultiplePickerHostProps> = React.createElement(PropertyFieldSPListMultiplePickerHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      context: this.context,
      selectedLists: this.selectedLists,
      baseTemplate: this.baseTemplate,
      orderBy: this.orderBy,
      includeHidden: this.includeHidden,
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
export function PropertyFieldSPListMultiplePicker(targetProperty: string, properties: IPropertyFieldSPListMultiplePickerProps): IPropertyPaneField<IPropertyFieldSPListMultiplePickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldSPListMultiplePickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      context: properties.context,
      selectedLists: properties.selectedLists,
      baseTemplate: properties.baseTemplate,
      orderBy: properties.orderBy,
      includeHidden: properties.includeHidden,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldSPListMultiplePicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSPListMultiplePickerBuilder(targetProperty, newProperties);
}
