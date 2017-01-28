/**
 * @file PropertyFieldMaskedInput.ts
 * Define a custom field of type PropertyFieldMaskedInput for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldMaskedInputHost, { IPropertyFieldMaskedInputHostProps } from './PropertyFieldMaskedInputHost';

/**
 * @interface
 * Public properties of the PropertyFieldMaskedInput custom field
 *
 */
export interface IPropertyFieldMaskedInputProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial value
   */
  initialValue?: string;
  /**
   * @var
   * Reg exp pattern
   */
  pattern: string;
  /**
   * @var
   * Placeholder
   */
  placeholder: string;
  /**
   * @var
   * Max length
   */
  maxLength: string;
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
   * Whether the property pane field is enabled or not.
   */
  disabled?: boolean;
}

/**
 * @interface
 * Private properties of the PropertyFieldMaskedInput custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldMaskedInput.
 *
 */
export interface IPropertyFieldMaskedInputPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  pattern: string;
  placeholder: string;
  maxLength: string;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  disabled?: boolean;
}

/**
 * @interface
 * Represents a PropertyFieldMaskedInput object
 *
 */
class PropertyFieldMaskedInputBuilder implements IPropertyPaneField<IPropertyFieldMaskedInputPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldMaskedInputPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private pattern: string;
  private placeholder: string;
  private maxLength: string;

  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldMaskedInputPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.pattern = _properties.pattern;
    this.placeholder = _properties.placeholder;
    this.maxLength = _properties.maxLength;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
  }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldMaskedInputHostProps> = React.createElement(PropertyFieldMaskedInputHost, {
      label: this.label,
      initialValue: this.initialValue,
      pattern: this.pattern,
      placeholder: this.placeholder,
      maxLength: this.maxLength,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
      disabled: this.disabled
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
export function PropertyFieldMaskedInput(targetProperty: string, properties: IPropertyFieldMaskedInputProps): IPropertyPaneField<IPropertyFieldMaskedInputPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldMaskedInputPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      pattern: properties.pattern,
      placeholder: properties.placeholder,
      maxLength: properties.maxLength,
      initialValue: properties.initialValue,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled
    };
    //Calles the PropertyFieldMaskedInput builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldMaskedInputBuilder(targetProperty, newProperties);
}


