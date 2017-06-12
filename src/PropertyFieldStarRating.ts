/**
 * @file PropertyFieldStarRating.ts
 * Define a custom field of type PropertyFieldStarRating for
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
import PropertyFieldStarRatingHost, { IPropertyFieldStarRatingHostProps } from './PropertyFieldStarRatingHost';

/**
 * @interface
 * Public properties of the PropertyFieldStarRating custom field
 *
 */
export interface IPropertyFieldStarRatingProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial value. Number of selected icon (`0` - none, `1` - first)
   */
  initialValue?: number;
  /**
   * @var
   * Number of icons in rating, default `5`
   */
  starCount?: number;
  /**
   * @var
   * Color of selected icons, default `#ffb400`
   */
  starColor?: string;
  /**
   * @var
   * Star size, default `24`
   */
  starSize?: number;
  /**
   * @var
   * Color of non-selected icons, default `#333`
   */
  emptyStarColor?: string;
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
   onGetErrorMessage?: (value: number) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldStarRating custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldStarRating.
 *
 */
export interface IPropertyFieldStarRatingPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: number;
  starCount?: number;
  starColor?: string;
  starSize?: number;
  emptyStarColor?: string;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  disabled?: boolean;
  onGetErrorMessage?: (value: number) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldStarRating object
 *
 */
class PropertyFieldStarRatingBuilder implements IPropertyPaneField<IPropertyFieldStarRatingPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldStarRatingPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: number;
  private starCount: number = 5;
  private starSize: number = 24;
  private starColor: string = '#ffb400';
  private emptyStarColor: string = '#333';
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: number) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldStarRatingPropsInternal) {
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
    if (_properties.starCount !== undefined)
      this.starCount = _properties.starCount;
    if (_properties.starColor !== undefined)
      this.starColor = _properties.starColor;
    if (_properties.emptyStarColor !== undefined)
      this.emptyStarColor = _properties.emptyStarColor;
    if (_properties.starSize !== undefined)
      this.starSize = _properties.starSize;
  }

  /**
   * @function
   * Renders the picker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldStarRatingHostProps> = React.createElement(PropertyFieldStarRatingHost, {
      label: this.label,
      initialValue: this.initialValue,
      starCount: this.starCount,
      starColor: this.starColor,
      emptyStarColor: this.emptyStarColor,
      starSize: this.starSize,
      targetProperty: this.targetProperty,
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
 * Helper method to create the customer field on the PropertyPane.
 * @param targetProperty - Target property the custom field is associated to.
 * @param properties - Strongly typed custom field properties.
 */
export function PropertyFieldStarRating(targetProperty: string, properties: IPropertyFieldStarRatingProps): IPropertyPaneField<IPropertyFieldStarRatingPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldStarRatingPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      starCount: properties.starCount,
      starColor: properties.starColor,
      starSize: properties.starSize,
      emptyStarColor: properties.emptyStarColor,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldStarRating builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldStarRatingBuilder(targetProperty, newProperties);
}


