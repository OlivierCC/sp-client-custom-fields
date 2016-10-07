/**
 * @file PropertyFieldSliderRange.ts
 * Define a custom field of type PropertyFieldSliderRange for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  IPropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-client-preview';
import PropertyFieldSliderRangeHost, { IPropertyFieldSliderRangeHostProps } from './PropertyFieldSliderRangeHost';
import ModuleLoader from '@microsoft/sp-module-loader';

/**
 * @interface
 * Public properties of the PropertyFieldSliderRange custom field
 *
 */
export interface IPropertyFieldSliderRangeProps {
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
   * @function
   * Defines a onPropertyChange function to raise when the selected Color changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Private properties of the PropertyFieldSliderRange custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSliderRange.
 *
 */
export interface IPropertyFieldSliderRangePropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Represents a PropertyFieldSliderRange object
 *
 */
class PropertyFieldSliderRangeBuilder implements IPropertyPaneField<IPropertyFieldSliderRangePropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSliderRangePropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private guid: string;
  private onPropertyChange: (propertyPath: string, newValue: any) => void;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSliderRangePropsInternal) {
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.render = this.render.bind(this);
    this.guid = this.getGuid();
  }

  private getGuid(): string {
    return this.s4() + this.s4() + '-' + this.s4() + '-' + this.s4() + '-' +
      this.s4() + '-' + this.s4() + this.s4() + this.s4();
  }

  private s4(): string {
      return Math.floor((1 + Math.random()) * 0x10000)
        .toString(16)
        .substring(1);
    }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {

    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldSliderRangeHostProps> = React.createElement(PropertyFieldSliderRangeHost, {
      label: this.label,
      initialValue: this.initialValue,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      guid: this.guid
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);

    var jQueryCdn = '//cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js';
    var jQueryUICdn = '//cdnjs.cloudflare.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js';
    var jQRangeSliderCdn = '//cdnjs.cloudflare.com/ajax/libs/jQRangeSlider/5.7.2/jQRangeSlider.min.js';

    ModuleLoader.loadScript(jQueryCdn, 'jQuery').then((jQuery: any): void => {
      ModuleLoader.loadScript(jQueryUICdn, 'jqueryui').then((jqueryui: any): void => {
        ModuleLoader.loadScript(jQRangeSliderCdn, 'jQRangeSlider').then((jQRangeSlider: any): void => {

        });
      });
    });
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
export function PropertyFieldSliderRange(targetProperty: string, properties: IPropertyFieldSliderRangeProps): IPropertyPaneField<IPropertyFieldSliderRangePropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldSliderRangePropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      onPropertyChange: properties.onPropertyChange,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldSliderRange builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSliderRangeBuilder(targetProperty, newProperties);
}


