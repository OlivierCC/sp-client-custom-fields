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
import { Async } from 'office-ui-fabric-react/lib/Utilities';

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
  onGetErrorMessage?: (value: string) => string | Promise<string>;
  deferredValidationTime?: number;
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
  private onGetErrorMessage: (value: string) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

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
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.deferredValidationTime);

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
      key: this.key,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime
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
              this.delayedValidate(value);
              /*if (this.onPropertyChange && value != null) {
                this.customProperties[this.targetProperty] = value;
                this.onPropertyChange(this.targetProperty, this.initialValue, value);
              }*/
              ($ as any)('#' + this.guid + '-min').html(ui.values[0]);
              ($ as any)('#' + this.guid + '-max').html(ui.values[1]);
            }.bind(this)
          });
      });
    });
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.onGetErrorMessage === null || this.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.initialValue, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.initialValue, value);
        ((document.getElementById(this.guid + '-errorMssg1')) as any).innerHTML = result;
        ((document.getElementById(this.guid + '-errorMssg2')) as any).innerHTML = result;
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.initialValue, value);
          ((document.getElementById(this.guid + '-errorMssg1')) as any).innerHTML = errorMessage;
          ((document.getElementById(this.guid + '-errorMssg2')) as any).innerHTML = errorMessage;
        });
      }
    }
    else {
      this.notifyAfterValidate(this.initialValue, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    if (this.onPropertyChange && newValue != null) {
      this.customProperties[this.targetProperty] = newValue;
      this.onPropertyChange(this.targetProperty, this.properties.initialValue, newValue);
    }
  }

  /**
   * @function
   * Disposes the current object
   */
  private dispose(elem: HTMLElement): void {
    if (this.async !== undefined)
      this.async.dispose();
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
 * Helper method to create the customer field on the PropertyPane.
 * @param targetProperty - Target property the custom field is associated to.
 * @param properties - Strongly typed custom field properties.
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
      key: properties.key,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldSliderRange builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSliderRangeBuilder(targetProperty, newProperties);
}


