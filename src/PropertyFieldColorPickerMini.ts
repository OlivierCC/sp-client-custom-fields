/**
 * @file PropertyFieldColorPickerMini.ts
 * Define a custom field of type PropertyFieldColorPickerMini for
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
import PropertyFieldColorPickerMiniHost, { IPropertyFieldColorPickerMiniHostProps } from './PropertyFieldColorPickerMiniHost';

/**
 * @interface
 * Public properties of the PropertyFieldColorPickerMini custom field
 *
 */
export interface IPropertyFieldColorPickerMiniProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial color
   */
  initialColor?: string;
  /**
   * Whether the property pane field is enabled or not. Default is `false`
   */
  disabled?: boolean;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected Color changed.
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
   onGetErrorMessage?: (value: string) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldColorPickerMini custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldColorPickerMini.
 *
 */
export interface IPropertyFieldColorPickerMiniPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialColor?: string;
  targetProperty: string;
  disabled?: boolean;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  onGetErrorMessage?: (value: string) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldColorPickerMini object
 *
 */
class PropertyFieldColorPickerMiniBuilder implements IPropertyPaneField<IPropertyFieldColorPickerMiniPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldColorPickerMiniPropsInternal;

  //Custom properties
  private label: string;
  private initialColor: string = '#FFFFFF';
  private disabled: boolean = false;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private onGetErrorMessage: (value: string) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldColorPickerMiniPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    if (_properties.initialColor !== undefined && _properties.initialColor != '')
      this.initialColor = _properties.initialColor;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    if (_properties.disabled !== undefined)
      this.disabled = _properties.disabled;
  }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldColorPickerMiniHostProps> = React.createElement(PropertyFieldColorPickerMiniHost, {
      label: this.label,
      initialColor: this.initialColor,
      targetProperty: this.targetProperty,
      disabled: this.disabled,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
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
 * Helper method to create a Color Picker on the PropertyPane.
 * @param targetProperty - Target property the Color picker is associated to.
 * @param properties - Strongly typed Color Picker properties.
 */
export function PropertyFieldColorPickerMini(targetProperty: string, properties: IPropertyFieldColorPickerMiniProps): IPropertyPaneField<IPropertyFieldColorPickerMiniPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldColorPickerMiniPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialColor: properties.initialColor,
      disabled: properties.disabled,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldColorPickerMini builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldColorPickerMiniBuilder(targetProperty, newProperties);
}


