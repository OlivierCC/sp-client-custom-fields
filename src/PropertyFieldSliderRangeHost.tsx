/**
 * @file PropertyFieldSliderRangeHost.tsx
 * Renders the controls for PropertyFieldSliderRange component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldSliderRangePropsInternal } from './PropertyFieldSliderRange';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldSliderRangeHost properties interface
 *
 */
export interface IPropertyFieldSliderRangeHostProps extends IPropertyFieldSliderRangePropsInternal {
  guid: string;
}


export interface IPropertyFieldSliderRangeHostState {
  scripLoaded: boolean;
}

/**
 * @class
 * Renders the controls for PropertyFieldSliderRange component
 */
export default class PropertyFieldSliderRangeHost extends React.Component<IPropertyFieldSliderRangeHostProps, IPropertyFieldSliderRangeHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldSliderRangeHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method

    //this.setState({scripLoaded: false});
  }


  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <div id={this.props.guid + '-editor'}></div>
      </div>
    );
  }
}
