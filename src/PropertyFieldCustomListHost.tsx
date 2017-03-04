/**
 * @file PropertyFieldCustomListHost.tsx
 * Renders the controls for PropertyFieldCustomList component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import styles from './PropertyFields.module.scss';
import { IPropertyFieldCustomListPropsInternal, ICustomListField, CustomListFieldType } from './PropertyFieldCustomList';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import PropertyFieldDatePickerHost from './PropertyFieldDatePickerHost';
import PropertyFieldDateTimePickerHost from './PropertyFieldDateTimePickerHost';
import PropertyFieldFontPickerHost from './PropertyFieldFontPickerHost';
import PropertyFieldFontSizePickerHost from './PropertyFieldFontSizePickerHost';
import PropertyFieldIconPickerHost from './PropertyFieldIconPickerHost';
import PropertyFieldColorPickerHost from './PropertyFieldColorPickerHost';
import PropertyFieldPasswordHost from './PropertyFieldPasswordHost';
import PropertyFieldPicturePickerHost from './PropertyFieldPicturePickerHost';
import PropertyFieldDocumentPickerHost from './PropertyFieldDocumentPickerHost';
import PropertyFieldSPListPickerHost from './PropertyFieldSPListPickerHost';
import PropertyFieldSPFolderPickerHost from './PropertyFieldSPFolderPickerHost';
import PropertyFieldPeoplePickerHost from './PropertyFieldPeoplePickerHost';
import GuidHelper from './GuidHelper';

import * as strings from 'sp-client-custom-fields/strings';

/**
 * @interface
 * PropertyFieldCustomListHost properties interface
 *
 */
export interface IPropertyFieldCustomListHostProps extends IPropertyFieldCustomListPropsInternal {
}

export interface IPropertyFieldCustomListHostState {
  data?: any[];
  openPanel?: boolean;
  openListView?: boolean;
  openListAdd?: boolean;
  openListEdit?: boolean;
  selectedIndex?: number;
  hoverColor?: string;
  deleteOpen?: boolean;
  editOpen?: boolean;
  mandatoryOpen?: boolean;
  missingField?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldCustomList component
 */
export default class PropertyFieldCustomListHost extends React.Component<IPropertyFieldCustomListHostProps, IPropertyFieldCustomListHostState> {

  private _key: string;

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldCustomListHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.saveWebPart = this.saveWebPart.bind(this);
    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClickAddItem = this.onClickAddItem.bind(this);
    this.onClickCancel = this.onClickCancel.bind(this);
    this.onClickAdd = this.onClickAdd.bind(this);
    this.onEnterHoverItem = this.onEnterHoverItem.bind(this);
    this.onLeaveHoverItem = this.onLeaveHoverItem.bind(this);
    this.onChangeSelectedItem = this.onChangeSelectedItem.bind(this);
    this.onClickDeleteItem = this.onClickDeleteItem.bind(this);
    this.onDismissDelete = this.onDismissDelete.bind(this);
    this.clickDelete = this.clickDelete.bind(this);
    this.onClickEdit = this.onClickEdit.bind(this);
    this.onClickUpdate = this.onClickUpdate.bind(this);
    this.onPropertyChange = this.onPropertyChange.bind(this);
    this.onPropertyChangeJson = this.onPropertyChangeJson.bind(this);
    this.onChangedCheckbox = this.onChangedCheckbox.bind(this);
    this.onCancel = this.onCancel.bind(this);
    this.onClickMoveUp = this.onClickMoveUp.bind(this);
    this.onClickMoveDown = this.onClickMoveDown.bind(this);
    this._key = GuidHelper.getGuid();

    this.state = {
      data: this.props.value != null ? this.props.value : [],
      openPanel: false,
      openListView: true,
      openListAdd: false,
      openListEdit: false,
      deleteOpen: false,
      editOpen: false,
      mandatoryOpen: false,
      missingField: ''
    };
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private saveWebPart(value: any[]): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && value != null) {
      this.props.properties[this.props.targetProperty] = value;
      this.props.onPropertyChange(this.props.targetProperty, [], value);
    }
  }

  private onOpenPanel(element?: any): void {
    this.state.openPanel = true;
    this.state.openListView = true;
    this.state.openListAdd = false;
    this.state.editOpen = false;
    this.state.mandatoryOpen = false;
    this.setState(this.state);
  }

  private onCancel(element?: any): void {
    this.state.openPanel = false;
    this.state.openListView = false;
    this.state.openListAdd = false;
    this.state.editOpen = false;
    this.state.mandatoryOpen = false;
    this.setState(this.state);
  }

  private onClickAddItem(element?: any): void {
    this.state.openListView = false;
    this.state.openListAdd = true;
    this.state.openListEdit = false;
    this.state.editOpen = false;
    this.state.mandatoryOpen = false;
    this.setState(this.state);
  }

  private onClickDeleteItem(element?: any): void {
    this.state.deleteOpen = true;
    this.setState(this.state);
  }

  private onClickCancel(): void {
    this.state.openListView = true;
    this.state.openListAdd = false;
    this.state.openListEdit = false;
    this.state.editOpen = false;
    this.state.mandatoryOpen = false;
    this.setState(this.state);
  }

  private onClickAdd(): void {
    var result = new Object();
    for (var i = 0; i < this.props.fields.length; i++) {
      if (this.props.fields[i] == null)
        continue;
      var ctrl = document.getElementById('input-' + this.props.fields[i].title);
      if (ctrl == null)
        continue;
      var str = ctrl['value'];
      if (str.length > 0 && (str[0] == '[' || str[0] == '{'))
        str = JSON.parse(str);

      if (this.props.fields[i].required === true && (str == null || str == '')) {
        this.state.mandatoryOpen = true;
        this.state.missingField = this.props.fields[i].title;
        this.setState(this.state);
        document.getElementById('input-' + this.props.fields[i].title).focus();
        return;
      }

      result[this.props.fields[i].title] = str;
    }
    this.state.data.push(result);
    this.setState(this.state);

    this.saveWebPart(this.state.data);

    this.onClickCancel();
  }

  private onChangeSelectedItem(element?: any): void {
    var index = element.currentTarget.value;
    this.state.selectedIndex = index;
    this.setState(this.state);
  }

  private onEnterHoverItem(element?: any): void {
    this.state.hoverColor = 'white';
    element.currentTarget.style.backgroundColor = '#F8F8F8';
  }
  private onLeaveHoverItem(element?: any): void {
    element.currentTarget.style.backgroundColor = this.state.hoverColor;
  }

  private onDismissDelete(element?: any): void {
    this.state.deleteOpen = false;
    this.setState(this.state);
  }

  private onClickMoveUp(element?: any): void {
     var indexToMove: number = Number(this.state.selectedIndex);
     if (indexToMove > 0) {
       var obj = this.state.data[indexToMove - 1];
       this.state.data[indexToMove - 1] = this.state.data[indexToMove];
       this.state.data[indexToMove] = obj;
       this.state.selectedIndex = indexToMove - 1;
       this.setState(this.state);
       this.saveWebPart(this.state.data);
     }
  }

  private onClickMoveDown(element?: any): void {
     var indexToMove: number = Number(this.state.selectedIndex);
     if (indexToMove < this.state.data.length - 1) {
       var dataRestore = this.state.data[indexToMove + 1];
       this.state.data[indexToMove + 1] = this.state.data[indexToMove];
       this.state.data[indexToMove] = dataRestore;
       this.state.selectedIndex = indexToMove + 1;
       this.setState(this.state);
       this.saveWebPart(this.state.data);
     }
  }

  private clickDelete(element?: any): void {
    var indexToDelete = this.state.selectedIndex;
    var newData: any[] = [];
    for (var i = 0; i < this.state.data.length; i++) {
      if (i != indexToDelete)
        newData.push(this.state.data[i]);
    }
    this.state.selectedIndex = -1;
    this.state.data = newData;
    this.state.selectedIndex = null;
    this.setState(this.state);
    this.onDismissDelete();
    this.saveWebPart(this.state.data);
  }

  private onClickEdit(element?: any): void {
    this.state.editOpen = true;
    this.state.openListView = false;
    this.setState(this.state);
  }

  private onClickUpdate(element?: any): void {

    var result = this.state.data[this.state.selectedIndex];
    for (var i = 0; i < this.props.fields.length; i++) {
      if (this.props.fields[i] == null)
        continue;
      var ctrl = document.getElementById('input-' + this.props.fields[i].title);
      if (ctrl == null)
        continue;
      var str = ctrl['value'];
      if (str.length > 0 && (str[0] == '[' || str[0] == '{'))
        str = JSON.parse(str);

      if (this.props.fields[i].required === true && (str == null || str == '')) {
        this.state.mandatoryOpen = true;
        this.state.missingField = this.props.fields[i].title;
        this.setState(this.state);
        document.getElementById('input-' + this.props.fields[i].title).focus();
        return;
      }

      result[this.props.fields[i].title] = str;
    }
    this.setState(this.state);
    this.saveWebPart(this.state.data);
    this.onClickCancel();
  }

  private onPropertyChange(targetProperty: string, oldValue?: any, newValue?: any): void {
    var input = document.getElementById(targetProperty);
    input['value'] = newValue;
  }

  private onPropertyChangeJson(targetProperty: string, oldValue?: any, newValue?: any): void {
    var input = document.getElementById(targetProperty);
    input['value'] = JSON.stringify(newValue);
  }

  private onChangedCheckbox(isChecked: boolean): void {

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


        <Dialog type={DialogType.close} isOpen={this.state.openPanel} title={this.props.headerText} onDismiss={this.onCancel}
                containerClassName={styles.msDialogMainCustom} isDarkOverlay={true} isBlocking={false}>

          <div style={{width: '630px', height: '500px', overflow: 'scroll'}}>

          { this.state.openListAdd === true ?
          <div>
               {this.props.fields != null ?
              <div>
              <div style={{marginBottom: '20px', backgroundColor: '#F4F4F4', width: '100%', paddingTop: '5px', paddingBottom: '5px'}}>
                <Button buttonType={ButtonType.hero} disabled={true} icon='Add'> &nbsp;{strings.CustomListAddItem}</Button>
                <Button buttonType={ButtonType.hero} onClick={this.onClickCancel} icon='Back'> &nbsp;{strings.CustomListBack}</Button>
              </div>
              { this.state.mandatoryOpen === true ?
                    <div className="ms-MessageBar">
                      <a name="anchorMessageBar"></a>
                      <div className="ms-MessageBar-content">
                        <div className="ms-MessageBar-icon">
                          <i className="ms-Icon ms-Icon--Error"></i>
                        </div>
                        <div className="ms-MessageBar-text">
                          {strings.CustomListFieldMissing.replace("{0}", this.state.missingField)}
                        </div>
                      </div>
                    </div>
                    : ''}
              <table className="ms-Table" cellSpacing="0" style={{marginTop: '30px', width: '100%', paddingRight:'10px'}}>
                  <tbody>
                      {
                        this.props.fields.map((value: ICustomListField, indexF: number) => {
                          return (
                            <tr key={this._key + '-customListTr1-' + indexF}>
                              <td><Label>{value.title}
                              {value.required === true ? ' (*)': ''}
                              </Label></td>
                              <td>
                                { value.type == CustomListFieldType.string ?
                                  <input id={'input-' + value.title} className='ms-TextField-field' style={{marginBottom: '8px'}}/>
                                : ''
                                }
                                { value.type == CustomListFieldType.number ?
                                  <input type="number" role="spinbutton" id={'input-' + value.title} aria-valuemax="99999" aria-valuemin="-999999" aria-valuenow="0" className='ms-TextField-field' style={{width: '100px', marginBottom: '8px'}} />
                                : ''
                                }
                                { value.type == CustomListFieldType.boolean ?
                                  <div  style={{marginBottom: '8px'}}>
                                    <input id={'input-' + value.title}  type="hidden" style={{visibility: 'hidden'}}/>
                                    <input type="radio" role="radio" aria-checked="false" name={'input-' + value.title} style={{width: '18px', height: '18px'}} value={'input-' + value.title} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = true;
                                        }
                                      }
                                    } /> <span style={{fontSize: '14px'}}>{strings.CustomListTrue}</span>
                                    <input type="radio" role="radio"  aria-checked="false" name={'input-' + value.title} style={{width: '18px', height: '18px'}} value={'input-' + value.title} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = false;
                                        }
                                      }
                                    } /> <span style={{fontSize: '14px'}}>{strings.CustomListFalse}</span>
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.date ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDatePickerHost key={'input-' + value.title} label="" properties={this.props.properties} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.dateTime ?
                                  <div>
                                    <input id={'input-' + value.title}  type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDateTimePickerHost key={'input-' + value.title} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.font ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontPickerHost key={'input-' + value.title} label=""  properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.fontSize ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontSizePickerHost key={'input-' + value.title} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.color ?
                                  <div>
                                    <input id={'input-' + value.title} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldColorPickerHost key={'input-' + value.title} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.icon ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldIconPickerHost key={'input-' + value.title} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.password ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPasswordHost key={'input-' + value.title} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.users ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPeoplePickerHost key={'input-' + value.title} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.list ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPListPickerHost key={'input-' + value.title} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.folder ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPFolderPickerHost key={'input-' + value.title} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.picture ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPicturePickerHost  key={'input-' + value.title} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.document ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDocumentPickerHost  key={'input-' + value.title} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                              </td>
                            </tr>
                          );
                        })
                      }
                  </tbody>
                  </table>
                  </div>
                : ''
               }
               <div style={{marginTop: '30px', marginBottom: '30px'}}>
                <Button buttonType={ButtonType.primary} onClick={this.onClickAdd}>{strings.CustomListOK}</Button>
                <Button buttonType={ButtonType.normal} onClick={this.onClickCancel}>{strings.CustomListCancel}</Button>
              </div>
          </div>
          : ''}

          { this.state.editOpen === true ?
          <div>
               {this.props.fields != null ?
                  <div>
                  <div style={{marginBottom: '20px', backgroundColor: '#F4F4F4', width: '100%', paddingTop: '5px', paddingBottom: '5px'}}>
                    <Button buttonType={ButtonType.hero} disabled={true} icon='Edit'> &nbsp;{strings.CustomListEdit}</Button>
                    <Button buttonType={ButtonType.hero} onClick={this.onClickCancel} icon='Back'> &nbsp;{strings.CustomListBack}</Button>
                  </div>
                  { this.state.mandatoryOpen === true ?
                    <div className="ms-MessageBar">
                      <a name="anchorMessageBar"></a>
                      <div className="ms-MessageBar-content">
                        <div className="ms-MessageBar-icon">
                          <i className="ms-Icon ms-Icon--Error"></i>
                        </div>
                        <div className="ms-MessageBar-text">
                          {strings.CustomListFieldMissing.replace("{0}", this.state.missingField)}
                        </div>
                      </div>
                    </div>
                    : ''}
                  <table className="ms-Table" cellSpacing="0" style={{marginTop: '30px', width: '100%', paddingRight:'10px'}}>
                  <tbody>
                      {
                        this.props.fields.map((value: ICustomListField, indexM: number) => {
                          return (
                            <tr key={this._key + '-customListTr2-' + indexM}>
                              <td><Label>{value.title}
                              {value.required === true ? ' (*)': ''}
                              </Label></td>
                              <td>
                                { value.type == CustomListFieldType.string ?
                                  <input id={'input-' + value.title} className='ms-TextField-field' style={{marginBottom: '8px'}} defaultValue={this.state.data[this.state.selectedIndex][value.title]} />
                                : ''
                                }
                                { value.type == CustomListFieldType.number ?
                                  <input type="number" role="spinbutton" id={'input-' + value.title} className='ms-TextField-field' defaultValue={this.state.data[this.state.selectedIndex][value.title]} aria-valuemax="99999" aria-valuemin="-999999" aria-valuenow={this.state.data[this.state.selectedIndex][value.title]} style={{width: '100px', marginBottom: '8px'}} />
                                : ''
                                }
                                { value.type == CustomListFieldType.boolean ?
                                  <div  style={{marginBottom: '8px'}}>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]} style={{visibility: 'hidden'}}/>
                                    <input type="radio" role="radio" name={'input-' + value.title} style={{width: '18px', height: '18px'}} value={'input-' + value.title} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = true;
                                        }
                                      }
                                    }
                                    defaultChecked={this.state.data[this.state.selectedIndex][value.title] == "true"}
                                    aria-checked={this.state.data[this.state.selectedIndex][value.title] == "true"}
                                    /> <span style={{fontSize: '14px'}}>{strings.CustomListTrue}</span>
                                    <input type="radio" role="radio" name={'input-' + value.title} style={{width: '18px', height: '18px'}} value={'input-' + value.title} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = false;
                                        }
                                      }
                                    }
                                    defaultChecked={this.state.data[this.state.selectedIndex][value.title] == "false"}
                                    aria-checked={this.state.data[this.state.selectedIndex][value.title] == "false"}
                                    /> <span style={{fontSize: '14px'}}>{strings.CustomListFalse}</span>
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.date ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDatePickerHost key={'input-' + value.title} properties={this.props.properties}  initialDate={this.state.data[this.state.selectedIndex][value.title]} label="" onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.dateTime ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDateTimePickerHost key={'input-' + value.title} properties={this.props.properties}  initialDate={this.state.data[this.state.selectedIndex][value.title]} label="" onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.font ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontPickerHost key={'input-' + value.title} label="" properties={this.props.properties}  initialValue={this.state.data[this.state.selectedIndex][value.title]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.fontSize ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontSizePickerHost  key={'input-' + value.title} properties={this.props.properties} label="" initialValue={this.state.data[this.state.selectedIndex][value.title]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.color ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldColorPickerHost key={'input-' + value.title} properties={this.props.properties}  label="" initialColor={this.state.data[this.state.selectedIndex][value.title]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.icon ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldIconPickerHost key={'input-' + value.title} properties={this.props.properties}  label="" initialValue={this.state.data[this.state.selectedIndex][value.title]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.password ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPasswordHost key={'input-' + value.title} properties={this.props.properties}  label="" initialValue={this.state.data[this.state.selectedIndex][value.title]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.users ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={JSON.stringify(this.state.data[this.state.selectedIndex][value.title])}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPeoplePickerHost key={'input-' + value.title} properties={this.props.properties}  label="" initialData={this.state.data[this.state.selectedIndex][value.title]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.list ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPListPickerHost properties={this.props.properties}  label="" selectedList={this.state.data[this.state.selectedIndex][value.title]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title} key={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.folder ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPFolderPickerHost key={'input-' + value.title} properties={this.props.properties}  label="" initialFolder={this.state.data[this.state.selectedIndex][value.title]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.picture ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPicturePickerHost key={'input-' + value.title} properties={this.props.properties}  label=""  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.document ?
                                  <div>
                                    <input id={'input-' + value.title} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.title]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDocumentPickerHost key={'input-' + value.title} properties={this.props.properties}  label=""  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.title}  />
                                  </div>
                                : ''
                                }
                              </td>
                            </tr>
                          );
                        })
                      }
                  </tbody>
                  </table>
                  </div>
                : ''
               }
               <div style={{marginTop: '30px', marginBottom: '30px'}}>
                <Button buttonType={ButtonType.primary} onClick={this.onClickUpdate}>{strings.CustomListOK}</Button>
                <Button buttonType={ButtonType.normal} onClick={this.onClickCancel}>{strings.CustomListCancel}</Button>
              </div>



          </div>
          : ''}

          { this.state.openListView === true ?
          <div>
              <div style={{marginBottom: '20px', backgroundColor: '#F4F4F4', width: '100%', paddingTop: '5px', paddingBottom: '5px'}}>
                <Button buttonType={ButtonType.hero} onClick={this.onClickAddItem} icon='Add'> &nbsp;{strings.CustomListAddItem}</Button>
                <Button buttonType={ButtonType.hero} onClick={this.onClickEdit} disabled={this.state.selectedIndex == null || this.state.selectedIndex < 0 ? true:false} icon='Edit'> &nbsp;{strings.CustomListEdit}</Button>
                <Button buttonType={ButtonType.hero} onClick={this.onClickDeleteItem} disabled={this.state.selectedIndex == null || this.state.selectedIndex < 0 ? true:false} icon='Delete'> &nbsp;{strings.CustomListDel}</Button>
                <Button buttonType={ButtonType.hero} onClick={this.onClickMoveUp} disabled={this.state.selectedIndex == null || this.state.selectedIndex < 0 ? true:false} icon='ChevronUp'> </Button>
                <Button buttonType={ButtonType.hero} onClick={this.onClickMoveDown} disabled={this.state.selectedIndex == null || this.state.selectedIndex < 0 ? true:false} icon='ChevronDown'> </Button>
              </div>
                 <Dialog type={DialogType.close} isOpen={this.state.deleteOpen} title={strings.CustomListConfirmDel}
                  onDismiss={this.onDismissDelete}  isDarkOverlay={false} isBlocking={true}>
                    <div>
                      <div>
                        <Label>{strings.CustomListConfirmDelMssg}</Label>
                      </div>
                      <div style={{paddingTop:'20px'}}>
                        <Button buttonType={ButtonType.primary} onClick={this.clickDelete}>{strings.CustomListYes}</Button>
                        <Button buttonType={ButtonType.normal} onClick={this.onDismissDelete}>{strings.CustomListNo}</Button>
                      </div>
                    </div>
                 </Dialog>

                {this.props.fields != null ?
                  <table className="ms-Table" cellPadding="4" cellSpacing="0" style={{width:'100%'}}>
                  <thead>
                    <tr>
                      <th style={{width: '35px', backgroundColor: '#F4F4F4', borderBottom: '1px', borderBottomColor: '#999999', borderBottomStyle: 'solid'}}></th>
                      {
                        this.props.fields.map((value: ICustomListField, indexJ: number) => {
                          if (value.hidden != true) {
                            return (
                              <th
                                key={this._key + '-customListTr3-' + indexJ}
                                style={{textAlign: 'left', backgroundColor: '#F4F4F4', borderBottom: '1px', borderBottomColor: '#999999', borderBottomStyle: 'solid'}}>
                                <Label style={{color: '#999999'}}>{value.title}</Label></th>
                            );
                          }
                        })
                      }
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {
                      this.state.data != null ?
                        this.state.data.map((value: any, index: number) => {
                          return (
                            <tr style={{backgroundColor: index == this.state.selectedIndex ? '#E8E8E8': ''}}
                              onMouseEnter={index != this.state.selectedIndex ? this.onEnterHoverItem : null}
                              onMouseLeave={index != this.state.selectedIndex ? this.onLeaveHoverItem : null}
                              key={this._key + '-customListTr4-' + index}
                              >
                              <td style={{paddingLeft: '6px', height:'40px'}}>
                                <div style={{float: 'left'}}>
                                <input id={"bulletRadio" + index} style={{width: '18px', height: '18px'}}
                                  type="radio" role="radio" name="radio1" onChange={this.onChangeSelectedItem}
                                  value={index.toString()} checked={index == this.state.selectedIndex ? true : false}
                                  aria-checked={index == this.state.selectedIndex ? true : false} />
                                <label htmlFor={"bulletRadio" + index}>
                                  <span className="ms-Label">
                                  </span>
                                </label>
                              </div>
                              </td>
                            {
                              this.props.fields.map((field: ICustomListField, indexI: number) => {
                                if (value != null && field != null && field.hidden != true) {
                                  return (
                                    <td key={this._key + '-customListTr5-' + indexI}><Label htmlFor={"bulletRadio" + index} >{value[field.title]}</Label></td>
                                  );
                                }
                              })
                            }
                            </tr>
                          );
                        })
                      : ''
                    }
                  </tbody>
                  </table>
                : '' }

          </div>
          : '' }

          </div>
        </Dialog>

        <Button disabled={this.props.disabled} onClick={this.onOpenPanel}>{this.props.headerText}</Button>

      </div>
    );
  }
}
