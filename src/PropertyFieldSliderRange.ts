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
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldSliderRangeHost, { IPropertyFieldSliderRangeHostProps } from './PropertyFieldSliderRangeHost';
import { SPComponentLoader } from '@microsoft/sp-loader';

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
   * @var
   * Disables the slider if set to true. Default is false.
   */
  disabled?: boolean;
  /**
   * @var
   * The minimum value of the slider. Default is 0.
   */
  min?: number;
  /**
   * @var
   * The maximum value of the slider. Default is 100.
   */
  max?: number;
  /**
   * @var
   * Default 1 - Determines the size or amount of each interval or step the
   * slider takes between the min and max. The full specified value range of the
   * slider (max - min) should be evenly divisible by the step.
   */
  step?: number;
  /**
   * @var
   * Determines whether the slider handles move horizontally (min on left,
   * max on right) or vertically (min on bottom, max on top).
   * Possible values: "horizontal", "vertical".
   */
  orientation?: string;
  /**
   * @var
   * Display the value on left & right of the slider or not
   */
  showValue?: boolean;
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
   * Initial value
   */
  key?: string;
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
  showValue?: boolean;
  guid: string;
  disabled?: boolean;
  min?: number;
  max?: number;
  step?: number;
  orientation?: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldSliderRange object
 *
 */
class PropertyFieldSliderRangeBuilder implements IPropertyPaneField<IPropertyFieldSliderRangePropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSliderRangePropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private guid: string;
  private disabled: boolean;
  private min: number;
  private max: number;
  private step: number;
  private orientation: string;
  private showValue: boolean;

  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSliderRangePropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.disabled = _properties.disabled;
    this.min = _properties.min;
    this.max = _properties.max;
    this.step = _properties.step;
    this.showValue = _properties.showValue;
    this.orientation = _properties.orientation;
    this.guid = _properties.guid;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;

    SPComponentLoader.loadCss('//cdnjs.cloudflare.com/ajax/libs/jqueryui/1.12.1/themes/smoothness/jquery-ui.css');
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
      disabled: this.disabled,
      min: this.min,
      max: this.max,
      step: this.step,
      orientation: this.orientation,
      showValue: this.showValue,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      guid: this.guid,
      properties: this.customProperties,
      key: this.key
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);

    var jQueryCdn = '//cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js';
    var jQueryUICdn = '//cdnjs.cloudflare.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js';

    SPComponentLoader.loadScript(jQueryCdn, { globalExportsName: '$' }).then(($: any): void => {
      SPComponentLoader.loadScript(jQueryUICdn, { globalExportsName: '$' }).then((jqueryui: any): void => {
          ($ as any)('#' + this.guid + '-slider').slider({
            range: true,
            min: this.min != null ? this.min : 0,
            max: this.max != null ? this.max : 100,
            step: this.step != null ? this.step : 1,
            disabled: this.disabled != null ? this.disabled : false,
            orientation: this.orientation != null ? this.orientation : 'horizontal',
            values: (this.initialValue != null && this.initialValue != '' && this.initialValue.split(",").length == 2) ? [ Number(this.initialValue.split(",")[0]), Number(this.initialValue.split(",")[1]) ] : [this.min, this.max],
            slide: function( event, ui ) {
              var value: string = ui.values[ 0 ] + "," + ui.values[ 1];
              if (this.onPropertyChange && value != null) {
                this.customProperties[this.targetProperty] = value;
                this.onPropertyChange(this.targetProperty, this.initialValue, value);
              }
              ($ as any)('#' + this.guid + '-min').html(ui.values[0]);
              ($ as any)('#' + this.guid + '-max').html(ui.values[1]);
            }.bind(this)
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


function s4(): string {
    return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);
}

function getGuid(): string {
    return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
      s4() + '-' + s4() + s4() + s4();
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
      disabled: properties.disabled,
      min: properties.min,
      max: properties.max,
      step: properties.step,
      showValue: properties.showValue,
      orientation: properties.orientation,
      guid: getGuid(),
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key
    };
    //Calles the PropertyFieldSliderRange builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSliderRangeBuilder(targetProperty, newProperties);
}


