/**
 * @file PropertyFieldCustomList.ts
 * Define a custom field of type PropertyFieldCustomList for
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
import PropertyFieldCustomListHost, { IPropertyFieldCustomListHostProps } from './PropertyFieldCustomListHost';
import { IWebPartContext } from '@microsoft/sp-webpart-base';

export enum CustomListFieldType {
  string = 0,
  number = 1,
  date = 2,
  boolean = 3,
  dateTime = 4,
  font = 5,
  fontSize = 6,
  color = 7,
  icon = 8,
  password = 9,
  picture = 10,
  document = 11,
  list = 12,
  users = 13,
  folder = 14
}

export interface ICustomListField {
  title: string;
  type: CustomListFieldType;
  required?: boolean;
  hidden?: boolean;
}

/**
 * @interface
 * Public properties of the PropertyFieldCustomList custom field
 *
 */
export interface IPropertyFieldCustomListProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Defines the Panel title
   */
  headerText: string;
  /**
   * @var
   * Defines the fields of the list
   */
  fields: ICustomListField[];
  /**
   * @var
   * Initial value
   */
  value?: any[];
  /**
   * @var
   * Parent web part context
   */
  context: IWebPartContext;
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
}

/**
 * @interface
 * Private properties of the PropertyFieldCustomList custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldCustomList.
 *
 */
export interface IPropertyFieldCustomListPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  fields: ICustomListField[];
  value?: any[];
  headerText: string;
  targetProperty: string;
  context: IWebPartContext;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldCustomList object
 *
 */
class PropertyFieldCustomListBuilder implements IPropertyPaneField<IPropertyFieldCustomListPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldCustomListPropsInternal;

  //Custom properties
  private label: string;
  private fields: ICustomListField[];
  private value: any[];
  private headerText: string;
  private context: IWebPartContext;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldCustomListPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.value = _properties.value;
    this.fields = _properties.fields;
    this.headerText = _properties.headerText;
    this.context = _properties.context;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
  }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldCustomListHostProps> = React.createElement(PropertyFieldCustomListHost, {
      label: this.label,
      value: this.value,
      headerText: this.headerText,
      fields: this.fields,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      context: this.context,
      properties: this.customProperties,
      key: this.targetProperty,
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
export function PropertyFieldCustomList(targetProperty: string, properties: IPropertyFieldCustomListProps): IPropertyPaneField<IPropertyFieldCustomListPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldCustomListPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      headerText: properties.headerText,
      value: properties.value,
      fields: properties.fields,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      context: properties.context,
      onDispose: null,
      onRender: null,
      key: targetProperty,
    };
    //Calles the PropertyFieldCustomList builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldCustomListBuilder(targetProperty, newProperties);
}


