/**
 * @file PropertyFieldTagPicker.ts
 * Define a custom field of type PropertyFieldTagPicker for
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
import PropertyFieldTagPickerHost, { IPropertyFieldTagPickerHostProps } from './PropertyFieldTagPickerHost';

/**
 * @interface
 * Generic Tag class
 */
export interface IPropertyFieldTag {
  /**
   * Tag's Key
   */
  key: string;
  /**
   * Tag's Name
   */
  name: string;
}

/**
 * @interface
 * Public properties of the PropertyFieldTagPicker custom field
 *
 */
export interface IPropertyFieldTagPickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Default Selected Tags
   */
  selectedTags?: IPropertyFieldTag[];
  /**
   * @var
   * Suggestions Header Text
   */
  suggestionsHeaderText: string;
  /**
   * @var
   * Text to display when no results found
   */
  noResultsFoundText: string;
  /**
   * @var
   * Text to display during loading
   */
  loadingText: string;
  /**
   * @var
   * List of tags
   */
  tags: IPropertyFieldTag[];
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
   onGetErrorMessage?: (value: IPropertyFieldTag[]) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldTagPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldTagPicker.
 *
 */
export interface IPropertyFieldTagPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  selectedTags?: IPropertyFieldTag[];
  suggestionsHeaderText: string;
  noResultsFoundText: string;
  loadingText: string;
  tags: IPropertyFieldTag[];
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  disabled?: boolean;
  onGetErrorMessage?: (value: IPropertyFieldTag[]) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldTagPicker object
 *
 */
class PropertyFieldTagPickerBuilder implements IPropertyPaneField<IPropertyFieldTagPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldTagPickerPropsInternal;

  //Custom properties
  private label: string;
  private selectedTags: IPropertyFieldTag[];
  private suggestionsHeaderText: string;
  private noResultsFoundText: string;
  private loadingText: string;
  private tags: IPropertyFieldTag[];
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: IPropertyFieldTag[]) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldTagPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.selectedTags = _properties.selectedTags;
    this.suggestionsHeaderText = _properties.suggestionsHeaderText;
    this.noResultsFoundText = _properties.noResultsFoundText;
    this.loadingText = _properties.loadingText;
    this.tags = _properties.tags;
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
  }

  /**
   * @function
   * Renders the picker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldTagPickerHostProps> = React.createElement(PropertyFieldTagPickerHost, {
      label: this.label,
      selectedTags: this.selectedTags,
      suggestionsHeaderText: this.suggestionsHeaderText,
      noResultsFoundText: this.noResultsFoundText,
      loadingText: this.loadingText,
      tags: this.tags,
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
export function PropertyFieldTagPicker(targetProperty: string, properties: IPropertyFieldTagPickerProps): IPropertyPaneField<IPropertyFieldTagPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldTagPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      selectedTags: properties.selectedTags,
      suggestionsHeaderText: properties.suggestionsHeaderText,
      noResultsFoundText: properties.noResultsFoundText,
      loadingText: properties.loadingText,
      tags: properties.tags,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldTagPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldTagPickerBuilder(targetProperty, newProperties);
}


