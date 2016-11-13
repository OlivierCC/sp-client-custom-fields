/**
 * @file PropertyFieldColorPickerHost.tsx
 * Renders the controls for PropertyFieldColorPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldColorPickerPropsInternal } from './PropertyFieldColorPicker';
import { ColorPicker } from 'office-ui-fabric-react/lib/ColorPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldColorPickerHost properties interface
 *
 */
export interface IPropertyFieldColorPickerHostProps extends IPropertyFieldColorPickerPropsInternal {
}

/**
 * @class
 * Renders the controls for PropertyFieldColorPicker component
 */
export default class PropertyFieldColorPickerHost extends React.Component<IPropertyFieldColorPickerHostProps, {}> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldColorPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onColorChanged = this.onColorChanged.bind(this);
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private onColorChanged(color: string): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && color != null) {
      this.props.properties[this.props.targetProperty] = color;
      this.props.onPropertyChange(this.props.targetProperty, this.props.initialColor, color);
    }
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    var defaultColor: string = '#FFFFFF';
    if (this.props.initialColor && this.props.initialColor != '')
      defaultColor = this.props.initialColor;
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <ColorPicker color={defaultColor} onColorChanged={this.onColorChanged}  />
      </div>
    );
  }
}