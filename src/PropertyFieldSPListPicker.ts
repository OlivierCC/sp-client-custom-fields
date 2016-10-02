/**
 * @file PropertyFieldSPListPicker.ts
 * Define a custom field of type PropertyFieldSPListPicker for
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
import PropertyFieldSPListPickerHost, { IPropertyFieldSPListPickerHostProps } from './PropertyFieldSPListPickerHost';


export enum PropertyFieldSPListPickerOrderBy {
  Id = 0,
  Title = 1
}

/**
 * @interface
 * Public properties of the PropertyFieldSPListPicker custom field
 *
 */
export interface IPropertyFieldSPListPickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  context: IWebPartContext;
  selectedList?: string;
  baseTemplate?: number;
  includeHidden?: boolean;
  orderBy?: PropertyFieldSPListPickerOrderBy;
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
 * Private properties of the PropertyFieldSPListPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSPListPicker.
 *
 */
export interface IPropertyFieldSPListPickerPropsInternal extends IPropertyFieldSPListPickerProps {
  label: string;
  targetProperty: string;
  context: IWebPartContext;
  selectedList?: string;
  baseTemplate?: number;
  orderBy?: PropertyFieldSPListPickerOrderBy;
  includeHidden?: boolean;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Represents a PropertyFieldSPListPicker object
 *
 */
class PropertyFieldSPListPickerBuilder implements IPropertyPaneField<IPropertyFieldSPListPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSPListPickerPropsInternal;

  //Custom properties label: string;
  private label: string;
  private context: IWebPartContext;
  private selectedList: string;
  private baseTemplate: number;
  private orderBy: PropertyFieldSPListPickerOrderBy;
  private includeHidden: boolean;

  public onPropertyChange(propertyPath: string, newValue: any): void {}

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSPListPickerPropsInternal) {
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.label = _properties.label;
    this.context = _properties.context;
    this.selectedList = _properties.selectedList;
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
    const element: React.ReactElement<IPropertyFieldSPListPickerHostProps> = React.createElement(PropertyFieldSPListPickerHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      context: this.context,
      selectedList: this.selectedList,
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
export function PropertyFieldSPListPicker(targetProperty: string, properties: IPropertyFieldSPListPickerProps): IPropertyPaneField<IPropertyFieldSPListPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldSPListPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      context: properties.context,
      selectedList: properties.selectedList,
      baseTemplate: properties.baseTemplate,
      orderBy: properties.orderBy,
      includeHidden: properties.includeHidden,
      onPropertyChange: properties.onPropertyChange,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldSPListPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSPListPickerBuilder(targetProperty, newProperties);
}
