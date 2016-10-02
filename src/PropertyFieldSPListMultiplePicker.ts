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
  IPropertyPaneFieldType,
  IWebPartContext
} from '@microsoft/sp-client-preview';
import PropertyFieldSPListMultiplePickerHost, { IPropertyFieldSPListMultiplePickerHostProps } from './PropertyFieldSPListMultiplePickerHost';

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
  orderBy?: PropertyFieldSPListMultiplePickerOrderBy;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected value changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, newValue: any): void;
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
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Represents a PropertyFieldSPListMultiplePicker object
 *
 */
class PropertyFieldSPListMultiplePickerBuilder implements IPropertyPaneField<IPropertyFieldSPListMultiplePickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSPListMultiplePickerPropsInternal;

  //Custom properties label: string;
  private label: string;
  private context: IWebPartContext;
  private selectedLists: string[];
  private baseTemplate: number;
  private orderBy: PropertyFieldSPListMultiplePickerOrderBy;
  private includeHidden: boolean;

  public onPropertyChange(propertyPath: string, newValue: any): void {}

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSPListMultiplePickerPropsInternal) {
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
      onPropertyChange: this.onPropertyChange
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
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldSPListMultiplePicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSPListMultiplePickerBuilder(targetProperty, newProperties);
}
