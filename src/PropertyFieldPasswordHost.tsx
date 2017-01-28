/**
 * @file PropertyFieldPasswordHost.tsx
 * Renders the controls for PropertyFieldPassword component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldPasswordPropsInternal } from './PropertyFieldPassword';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldPasswordHost properties interface
 *
 */
export interface IPropertyFieldPasswordHostProps extends IPropertyFieldPasswordPropsInternal {
}

/**
 * @class
 * Renders the controls for PropertyFieldPassword component
 */
export default class PropertyFieldPasswordHost extends React.Component<IPropertyFieldPasswordHostProps, {}> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldPasswordHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private onValueChanged(element: any): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && element != null) {
      this.props.properties[this.props.targetProperty] = element.currentTarget.value;
      this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, element.currentTarget.value);
    }
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <input disabled={this.props.disabled} label={this.props.label} type="password" value={this.props.initialValue} className="ms-TextField-field"
          onChange={this.onValueChanged}
          />
      </div>
    );
  }
}
