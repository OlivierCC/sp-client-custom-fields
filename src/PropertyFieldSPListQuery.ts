/**
 * @file PropertyFieldSPListQuery.ts
 * Define a custom field of type PropertyFieldSPListQuery for
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
import PropertyFieldSPListQueryHost, { IPropertyFieldSPListQueryHostProps } from './PropertyFieldSPListQueryHost';


export enum PropertyFieldSPListQueryOrderBy {
  Id = 0,
  Title = 1
}

/**
 * @interface
 * Public properties of the PropertyFieldSPListQuery custom field
 *
 */
export interface IPropertyFieldSPListQueryProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  context: IWebPartContext;
  query?: string;
  selectedList?: string;
  baseTemplate?: number;
  includeHidden?: boolean;
  orderBy?: PropertyFieldSPListQueryOrderBy;
  showOrderBy?: boolean;
  showMax?: boolean;
  showFilters?: boolean;
  max?: number;
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
 * Private properties of the PropertyFieldSPListQuery custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSPListQuery.
 *
 */
export interface IPropertyFieldSPListQueryPropsInternal extends IPropertyFieldSPListQueryProps {
  label: string;
  targetProperty: string;
  context: IWebPartContext;
  query?: string;
  selectedList?: string;
  baseTemplate?: number;
  orderBy?: PropertyFieldSPListQueryOrderBy;
  includeHidden?: boolean;
  showOrderBy?: boolean;
  showMax?: boolean;
  showFilters?: boolean;
  max?: number;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Represents a PropertyFieldSPListQuery object
 *
 */
class PropertyFieldSPListQueryBuilder implements IPropertyPaneField<IPropertyFieldSPListQueryPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSPListQueryPropsInternal;

  //Custom properties label: string;
  private label: string;
  private context: IWebPartContext;
  private query: string;
  private selectedList: string;
  private baseTemplate: number;
  private orderBy: PropertyFieldSPListQueryOrderBy;
  private includeHidden: boolean;
  private showOrderBy: boolean;
  private showMax: boolean;
  private showFilters: boolean;
  private max: number;
  public onPropertyChange(propertyPath: string, newValue: any): void {}

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSPListQueryPropsInternal) {
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.label = _properties.label;
    this.context = _properties.context;
    this.query = _properties.query;
    this.selectedList = _properties.selectedList;
    this.baseTemplate = _properties.baseTemplate;
    this.orderBy = _properties.orderBy;
    this.includeHidden = _properties.includeHidden;
    this.showOrderBy = _properties.showOrderBy;
    this.showMax = _properties.showMax;
    this.showFilters = _properties.showFilters;
    this.max = _properties.max;
    this.onPropertyChange = _properties.onPropertyChange;
  }

  /**
   * @function
   * Renders the SPListPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldSPListQueryHostProps> = React.createElement(PropertyFieldSPListQueryHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      context: this.context,
      query: this.query,
      selectedList: this.selectedList,
      baseTemplate: this.baseTemplate,
      orderBy: this.orderBy,
      includeHidden: this.includeHidden,
      showOrderBy: this.showOrderBy,
      showMax: this.showMax,
      showFilters: this.showFilters,
      max: this.max,
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
export function PropertyFieldSPListQuery(targetProperty: string, properties: IPropertyFieldSPListQueryProps): IPropertyPaneField<IPropertyFieldSPListQueryPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldSPListQueryPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      context: properties.context,
      query: properties.query,
      selectedList: properties.selectedList,
      baseTemplate: properties.baseTemplate,
      orderBy: properties.orderBy,
      includeHidden: properties.includeHidden,
      showOrderBy: properties.showOrderBy,
      showMax: properties.showMax,
      showFilters: properties.showFilters,
      max: properties.max,
      onPropertyChange: properties.onPropertyChange,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldSPListQuery builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSPListQueryBuilder(targetProperty, newProperties);
}
