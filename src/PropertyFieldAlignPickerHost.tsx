/**
 * @file PropertyFieldAlignPickerHost.tsx
 * Renders the controls for PropertyFieldAlignPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldAlignPickerPropsInternal } from './PropertyFieldAlignPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldAlignPickerHost properties interface
 *
 */
export interface IPropertyFieldAlignPickerHostProps extends IPropertyFieldAlignPickerPropsInternal {
}

export interface IPropertyFieldAlignPickerHostState {
  mode?: string;
  overList?: boolean;
  overTiles?: boolean;
  overRight?: boolean;
}

/**
 * @class
 * Renders the controls for PropertyFieldAlignPicker component
 */
export default class PropertyFieldAlignPickerHost extends React.Component<IPropertyFieldAlignPickerHostProps, IPropertyFieldAlignPickerHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldAlignPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
    this.onClickBullets = this.onClickBullets.bind(this);
    this.onClickTiles = this.onClickTiles.bind(this);
    this.onClickRight = this.onClickRight.bind(this);
    this.mouseListEnterDropDown = this.mouseListEnterDropDown.bind(this);
    this.mouseListLeaveDropDown = this.mouseListLeaveDropDown.bind(this);
    this.mouseTilesEnterDropDown = this.mouseTilesEnterDropDown.bind(this);
    this.mouseTilesLeaveDropDown = this.mouseTilesLeaveDropDown.bind(this);
    this.mouseRightEnterDropDown = this.mouseRightEnterDropDown.bind(this);
    this.mouseRightLeaveDropDown = this.mouseRightLeaveDropDown.bind(this);

    this.state = {
      mode: this.props.initialValue != null && this.props.initialValue != '' ? this.props.initialValue : '',
      overList: false, overTiles: false, overRight: false
    };
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private onValueChanged(element: any, previous: string, value: string): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChanged && element != null) {
      this.props.properties[this.props.targetProperty] = value;
      this.props.onPropertyChanged(this.props.targetProperty, previous, value);
    }
  }

  private onClickBullets(element?: any) {
    var previous = this.state.mode;
    this.state.mode = 'left';
    this.setState(this.state);
    this.onValueChanged(this, previous, this.state.mode);
  }

  private onClickTiles(element?: any) {
    var previous = this.state.mode;
    this.state.mode = 'center';
    this.setState(this.state);
    this.onValueChanged(this, previous, this.state.mode);
  }

  private onClickRight(element?: any) {
    var previous = this.state.mode;
    this.state.mode = 'right';
    this.setState(this.state);
    this.onValueChanged(this, previous, this.state.mode);
  }

  private mouseListEnterDropDown() {
    this.state.overList = true;
    this.setState(this.state);
  }

  private mouseListLeaveDropDown() {
    this.state.overList = false;
    this.setState(this.state);
  }

  private mouseTilesEnterDropDown() {
    this.state.overTiles = true;
    this.setState(this.state);
  }

  private mouseTilesLeaveDropDown() {
    this.state.overTiles = false;
    this.setState(this.state);
  }

  private mouseRightEnterDropDown() {
    this.state.overRight = true;
    this.setState(this.state);
  }

  private mouseRightLeaveDropDown() {
    this.state.overRight = false;
    this.setState(this.state);
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var backgroundTiles = this.state.overTiles ? '#DFDFDF': '';
    var backgroundLists = this.state.overList ? '#DFDFDF': '';
    var backgroundRight = this.state.overRight ? '#DFDFDF': '';
    if (this.state.mode == 'left')
      backgroundLists = '#EEEEEE';
    if (this.state.mode == 'center')
      backgroundTiles = '#EEEEEE';
    if (this.state.mode == 'right')
      backgroundRight = '#EEEEEE';

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>

        <div style={{display: 'inline-flex'}}>
          <div style={{cursor: 'pointer', width: '70px', marginRight: '30px', backgroundColor: backgroundLists}}
            onMouseEnter={this.mouseListEnterDropDown} onMouseLeave={this.mouseListLeaveDropDown}>
            <div style={{float: 'left'}}>

              <input id="bulletRadio" className=""
                onChange={this.onClickBullets} type="radio" name="radio1"
                defaultChecked={this.state.mode == "left" ? true : false}
                value="left"  style={{cursor: 'pointer', width: '18px', height: '18px'}}/>
              <label htmlFor="bulletRadio" className="">
                <span className="ms-Label">
                  <i className='ms-Icon ms-Icon--AlignLeft' aria-hidden="true" style={{cursor: 'pointer',fontSize:'32px', paddingLeft: '30px', color: '#808080'}}></i>
                </span>
              </label>
            </div>
          </div>
          <div style={{cursor: 'pointer', width: '70px', marginRight: '30px', backgroundColor: backgroundTiles}}
            onMouseEnter={this.mouseTilesEnterDropDown} onMouseLeave={this.mouseTilesLeaveDropDown}>
            <div style={{float: 'left'}}>
              <input id="tilesRadio" className=""
               onChange={this.onClickTiles} type="radio" name="radio1"
               defaultChecked={this.state.mode == "center" ? true : false}
               value="center"  style={{cursor: 'pointer', width: '18px', height: '18px'}}/>
              <label htmlFor="tilesRadio" className="">
                <span className="ms-Label">
                  <i className='ms-Icon ms-Icon--AlignCenter' aria-hidden="true" style={{cursor: 'pointer',fontSize:'32px', paddingLeft: '30px', color: '#808080'}}></i>
                </span>
              </label>
            </div>
          </div>
          <div style={{cursor: 'pointer', width: '70px', marginRight: '30px', backgroundColor: backgroundRight}}
            onMouseEnter={this.mouseRightEnterDropDown} onMouseLeave={this.mouseRightLeaveDropDown}>
            <div style={{float: 'left'}}>
              <input id="rightRadio" className=""
               onChange={this.onClickRight} type="radio" name="radio1"
               defaultChecked={this.state.mode == "right" ? true : false}
               value="right"  style={{cursor: 'pointer', width: '18px', height: '18px'}}/>
              <label htmlFor="rightRadio" className="">
                <span className="ms-Label">
                  <i className='ms-Icon ms-Icon--AlignRight' aria-hidden="true" style={{cursor: 'pointer',fontSize:'32px', paddingLeft: '30px', color: '#808080'}}></i>
                </span>
              </label>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
