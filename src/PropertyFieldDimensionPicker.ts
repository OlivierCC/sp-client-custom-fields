/**
 * @file PropertyFieldDimensionPicker.ts
 * Define a custom field of type PropertyFieldDimensionPicker for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldDimensionPickerHost, { IPropertyFieldDimensionPickerHostProps } from './PropertyFieldDimensionPickerHost';

/**
 * @interface
 * Defines a Dimension object for the PropertyFieldDimensionPicker
 *
 */
export interface IPropertyFieldDimension {
  width?: string;
  height?: string;
}

/**
 * @interface
 * Public properties of the PropertyFieldDimensionPicker custom field
 *
 */
export interface IPropertyFieldDimensionPickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial value
   */
  initialValue?: IPropertyFieldDimension;
  /**
   * @var
   * Whether the aspect ratio is checked or not by default. Default value is true.
   */
  preserveRatio?: boolean;
  /**
   * @var
   * Whether the aspect ratio checkbox is available or not. Default value is true.
   */
  preserveRatioEnabled?: boolean;
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
   onGetErrorMessage?: (value: IPropertyFieldDimension) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldDimensionPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldDimensionPicker.
 *
 */
export interface IPropertyFieldDimensionPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: IPropertyFieldDimension;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  disabled?: boolean;
  onGetErrorMessage?: (value: IPropertyFieldDimension) => string | Promise<string>;
  deferredValidationTime?: number;
  preserveRatio?: boolean;
  preserveRatioEnabled?: boolean;
}

/**
 * @interface
 * Represents a PropertyFieldDimensionPicker object
 *
 */
class PropertyFieldDimensionPickerBuilder implements IPropertyPaneField<IPropertyFieldDimensionPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldDimensionPickerPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: IPropertyFieldDimension;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: IPropertyFieldDimension) => string | Promise<string>;
  private deferredValidationTime: number = 200;
  private preserveRatio: boolean = true;
  private preserveRatioEnabled: boolean = true;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldDimensionPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    if (_properties.preserveRatio === false)
      this.preserveRatio = _properties.preserveRatio;
    if (_properties.preserveRatioEnabled === false)
      this.preserveRatioEnabled = _properties.preserveRatioEnabled;
  }

  /**
   * @function
   * Renders the Picker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldDimensionPickerHostProps> = React.createElement(PropertyFieldDimensionPickerHost, {
      label: this.label,
      initialValue: this.initialValue,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
      disabled: this.disabled,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime,
      preserveRatio: this.preserveRatio,
      preserveRatioEnabled: this.preserveRatioEnabled
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
 * Helper method to create the customer field on the PropertyPane.
 * @param targetProperty - Target property the custom field is associated to.
 * @param properties - Strongly typed custom field properties.
 */
export function PropertyFieldDimensionPicker(targetProperty: string, properties: IPropertyFieldDimensionPickerProps): IPropertyPaneField<IPropertyFieldDimensionPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldDimensionPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime,
      preserveRatio: properties.preserveRatio,
      preserveRatioEnabled: properties.preserveRatioEnabled
    };
    //Calls the PropertyFieldDimensionPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldDimensionPickerBuilder(targetProperty, newProperties);
}


