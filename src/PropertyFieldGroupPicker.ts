/**
 * @file PropertyFieldGroupPicker.ts
 * Define a custom field of type PropertyFieldGroupPicker for
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
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldGroupPickerHost, { IPropertyFieldGroupPickerHostProps } from './PropertyFieldGroupPickerHost';
import { IWebPartContext} from '@microsoft/sp-webpart-base';

/**
 * @enum
 * Group type
 */
export enum IGroupType {
  /**
   * SharePoint Group
   */
  SharePoint = 0,
  /**
   * Security Group
   */
  Security = 1
}

/**
 * @interface
 * Defines a Group object for the PropertyFieldGroupPicker
 *
 */
export interface IPropertyFieldGroup {
  /**
   * @var
   * Group's ID
   */
  id: string;
  /**
   * @var
   * Group's full name
   */
  fullName: string;
  /**
   * @var
   * Group's login
   */
  login: string;
  /**
   * @var
   * Group's description
   */
  description: string;
}

/**
 * @interface
 * Public properties of the PropertyFieldGroupPicker custom field
 *
 */
export interface IPropertyFieldGroupPickerProps {
  /**
   * @var
   * Property field label
   */
  label: string;
  /**
   * @var
   * Web Part context
   */
  context: IWebPartContext;
  /**
   * @var
   * Intial data to load in the people picker (optional)
   */
  initialData?: IPropertyFieldGroup[];
  /**
   * @var
   * Defines if the People Picker allows to select duplicated users (optional). Default is `false`
   */
  allowDuplicate?: boolean;
  /**
   * @var
   * Defines the groups type to request. Can be a SharePoint group, or a security group.
   */
  groupType: IGroupType;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected value changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  /**
   * @function
   * This API is called to render the web part.
   * Normally this function must be always defined with the 'this.render.bind(this)'
   * method of the web part object.
   */
  render(): void;
  /**
   * This property is used to indicate the web part's PropertyPane interaction mode: Reactive or NonReactive.
   * The default behaviour is Reactive.
   */
  disableReactivePropertyChanges?: boolean;
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
   onGetErrorMessage?: (value: IPropertyFieldGroup[]) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldGroupPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldGroupPicker.
 *
 */
export interface IPropertyFieldGroupPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  targetProperty: string;
  context: IWebPartContext;
  initialData?: IPropertyFieldGroup[];
  allowDuplicate?: boolean;
  groupType: IGroupType;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  render(): void;
  disableReactivePropertyChanges?: boolean;
  properties: any;
  onGetErrorMessage?: (value: IPropertyFieldGroup[]) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldGroupPicker object
 *
 */
class PropertyFieldGroupPickerBuilder implements IPropertyPaneField<IPropertyFieldGroupPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldGroupPickerPropsInternal;

  //Custom properties
  private label: string;
  private context: IWebPartContext;
  private initialData: IPropertyFieldGroup[];
  private allowDuplicate: boolean = false;
  private groupType: IGroupType;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private onGetErrorMessage: (value: IPropertyFieldGroup[]) => string | Promise<string>;
  private deferredValidationTime: number = 200;
  private renderWebPart: () => void;
  private disableReactivePropertyChanges: boolean = false;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldGroupPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.label = _properties.label;
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.context = _properties.context;
    this.initialData = _properties.initialData;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    this.groupType = _properties.groupType;
    if (_properties.allowDuplicate !== undefined)
      this.allowDuplicate = _properties.allowDuplicate;
    this.renderWebPart = _properties.render;
    if (_properties.disableReactivePropertyChanges !== undefined && _properties.disableReactivePropertyChanges != null)
      this.disableReactivePropertyChanges = _properties.disableReactivePropertyChanges;
  }

  /**
   * @function
   * Renders the PeoplePicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldGroupPickerHostProps> = React.createElement(PropertyFieldGroupPickerHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      initialData: this.initialData,
      allowDuplicate: this.allowDuplicate,
      groupType: this.groupType,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      context: this.context,
      properties: this.customProperties,
      key: this.key,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime,
      render: this.renderWebPart,
      disableReactivePropertyChanges: this.disableReactivePropertyChanges
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
 * Helper method to create a People Picker on the PropertyPane.
 * @param targetProperty - Target property the people picker is associated to.
 * @param properties - Strongly typed people Picker properties.
 */
export function PropertyFieldGroupPicker(targetProperty: string, properties: IPropertyFieldGroupPickerProps): IPropertyPaneField<IPropertyFieldGroupPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldGroupPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      onPropertyChange: properties.onPropertyChange,
      context: properties.context,
      initialData: properties.initialData,
      allowDuplicate: properties.allowDuplicate,
      groupType: properties.groupType,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime,
      render: properties.render,
      disableReactivePropertyChanges: properties.disableReactivePropertyChanges
    };
    //Calls the PropertyFieldGroupPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldGroupPickerBuilder(targetProperty, newProperties);
}


