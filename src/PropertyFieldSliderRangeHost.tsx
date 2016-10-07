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
        <table style={{paddingTop: '8px', paddingBottom: '10px', width:"100%"}} cellpadding="0" cellspacing="10">
        { this.props.showValue == false ?
            <tr><td width="100%"><div id={this.props.guid + '-slider'}></div></td></tr>
          :
            this.props.orientation == 'vertical' ?
              <tr>
                <td width="100%">
                  <div className="ms-Label" style={{marginBottom:'8px'}} id={this.props.guid + '-max'}>{(this.props.initialValue != null && this.props.initialValue != '' && this.props.initialValue.split(",").length == 2) ? this.props.initialValue.split(",")[1] : '0' }</div>
                  <div id={this.props.guid + '-slider'}></div>
                  <div className="ms-Label" style={{marginTop:'8px'}} id={this.props.guid + '-min'}>{(this.props.initialValue != null && this.props.initialValue != '' && this.props.initialValue.split(",").length == 2) ? this.props.initialValue.split(",")[0] : '0' }</div>
                </td>
              </tr>
            :
              <tr>
                <td width="35"><div className="ms-Label" id={this.props.guid + '-min'}>{(this.props.initialValue != null && this.props.initialValue != '' && this.props.initialValue.split(",").length == 2) ? this.props.initialValue.split(",")[0] : '0' }</div></td>
                <td width="220"><div id={this.props.guid + '-slider'}></div></td>
                <td width="35" style={{textAlign: 'right'}}><div className="ms-Label" id={this.props.guid + '-max'}>{(this.props.initialValue != null && this.props.initialValue != '' && this.props.initialValue.split(",").length == 2) ? this.props.initialValue.split(",")[1] : '0' }</div></td>
              </tr>
        }
        </table>
      </div>
    );
  }
}
