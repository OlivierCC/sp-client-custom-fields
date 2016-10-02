/**
 * @file PropertyFieldDisplayModeHost.tsx
 * Renders the controls for PropertyFieldDisplayMode component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDisplayModePropsInternal } from './PropertyFieldDisplayMode';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldDisplayModeHost properties interface
 *
 */
export interface IPropertyFieldDisplayModeHostProps extends IPropertyFieldDisplayModePropsInternal {
}

export interface IPropertyFieldDisplayModeHostState {
  mode?: string;
  overList?: boolean;
  overTiles?: boolean;
}

/**
 * @class
 * Renders the controls for PropertyFieldDisplayMode component
 */
export default class PropertyFieldDisplayModeHost extends React.Component<IPropertyFieldDisplayModeHostProps, IPropertyFieldDisplayModeHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldDisplayModeHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
    this.onClickBullets = this.onClickBullets.bind(this);
    this.onClickTiles = this.onClickTiles.bind(this);
    this.mouseListEnterDropDown = this.mouseListEnterDropDown.bind(this);
    this.mouseListLeaveDropDown = this.mouseListLeaveDropDown.bind(this);
    this.mouseTilesEnterDropDown = this.mouseTilesEnterDropDown.bind(this);
    this.mouseTilesLeaveDropDown = this.mouseTilesLeaveDropDown.bind(this);

    this.state = {
      mode: this.props.initialValue != null && this.props.initialValue != '' ? this.props.initialValue : '',
      overList: false, overTiles: false
    };
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private onValueChanged(element: any, value: string): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && element != null) {
      this.props.onPropertyChange(this.props.targetProperty, value);
    }
  }

  private onClickBullets(element?: any) {
    this.state.mode = 'list';
    this.setState(this.state);
    this.onValueChanged(this, this.state.mode);
  }

  private onClickTiles(element?: any) {
    this.state.mode = 'tiles';
    this.setState(this.state);
    this.onValueChanged(this, this.state.mode);
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

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var backgroundTiles = this.state.overTiles ? '#DFDFDF': '';
    var backgroundLists = this.state.overList ? '#DFDFDF': '';
    if (this.state.mode == 'list')
      backgroundLists = '#EEEEEE';
    if (this.state.mode == 'tiles')
      backgroundTiles = '#EEEEEE';

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>

        <div style={{display: 'inline-flex'}}>
          <div style={{cursor: 'pointer', width: '100px', marginRight: '30px', paddingLeft:'8px', backgroundColor: backgroundLists}}
            onMouseEnter={this.mouseListEnterDropDown} onMouseLeave={this.mouseListLeaveDropDown}>
            <div style={{float: 'left'}}>

              <input id="bulletRadio" className=""
                onChange={this.onClickBullets} type="radio" name="radio1"
                defaultChecked={this.state.mode == "list" ? true : false}
                value="list"  style={{cursor: 'pointer', width: '18px', height: '18px'}}/>
              <label htmlFor="bulletRadio" className="">
                <span className="ms-Label">
                  <i className='ms-Icon ms-Icon--List' aria-hidden="true" style={{cursor: 'pointer',fontSize:'60px', paddingLeft: '30px', color: '#808080'}}></i>
                </span>
              </label>
            </div>
          </div>
          <div style={{cursor: 'pointer', width: '100px', marginRight: '30px', paddingLeft:'8px', backgroundColor: backgroundTiles}}
            onMouseEnter={this.mouseTilesEnterDropDown} onMouseLeave={this.mouseTilesLeaveDropDown}>
            <div style={{float: 'left'}}>
              <input id="tilesRadio" className=""
               onChange={this.onClickTiles} type="radio" name="radio1"
               defaultChecked={this.state.mode == "tiles" ? true : false}
               value="tiles"  style={{cursor: 'pointer', width: '18px', height: '18px'}}/>
              <label htmlFor="tilesRadio" className="">
                <span className="ms-Label">
                  <i className='ms-Icon ms-Icon--Tiles' aria-hidden="true" style={{cursor: 'pointer',fontSize:'48px', paddingLeft: '30px', color: '#808080'}}></i>
                </span>
              </label>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
