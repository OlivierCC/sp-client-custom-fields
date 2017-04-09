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
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import {
  CheckboxVisibility,
  ConstrainMode,
  DetailsList,
  DetailsListLayoutMode as LayoutMode,
  SelectionMode,
  Selection,
  buildColumns
} from 'office-ui-fabric-react/lib/DetailsList';
import PropertyFieldDatePickerHost from './PropertyFieldDatePickerHost';
import PropertyFieldDateTimePickerHost from './PropertyFieldDateTimePickerHost';
import PropertyFieldFontPickerHost from './PropertyFieldFontPickerHost';
import PropertyFieldFontSizePickerHost from './PropertyFieldFontSizePickerHost';
import PropertyFieldIconPickerHost from './PropertyFieldIconPickerHost';
import PropertyFieldColorPickerHost from './PropertyFieldColorPickerHost';
import PropertyFieldColorPickerMiniHost from './PropertyFieldColorPickerMiniHost';
import PropertyFieldPasswordHost from './PropertyFieldPasswordHost';
import PropertyFieldPicturePickerHost from './PropertyFieldPicturePickerHost';
import PropertyFieldDocumentPickerHost from './PropertyFieldDocumentPickerHost';
import PropertyFieldSPListPickerHost from './PropertyFieldSPListPickerHost';
import PropertyFieldSPFolderPickerHost from './PropertyFieldSPFolderPickerHost';
import PropertyFieldPeoplePickerHost from './PropertyFieldPeoplePickerHost';
import PropertyFieldStarRatingHost from './PropertyFieldStarRatingHost';
import PropertyFieldGroupPickerHost from './PropertyFieldGroupPickerHost';
import { IGroupType } from './PropertyFieldGroupPicker';
import PropertyFieldOfficeVideoPickerHost from './PropertyFieldOfficeVideoPickerHost';
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
  items: any[];
  columns: any[];
  listKey: string;
  selection: Selection;
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
    this.onClickDeleteItem = this.onClickDeleteItem.bind(this);
    this.onDismissDelete = this.onDismissDelete.bind(this);
    this.clickDelete = this.clickDelete.bind(this);
    this.onClickEdit = this.onClickEdit.bind(this);
    this.onClickUpdate = this.onClickUpdate.bind(this);
    this.onPropertyChange = this.onPropertyChange.bind(this);
    this.onPropertyChangeJson = this.onPropertyChangeJson.bind(this);
    this.onCancel = this.onCancel.bind(this);
    this.onClickMoveUp = this.onClickMoveUp.bind(this);
    this.onClickMoveDown = this.onClickMoveDown.bind(this);
    this.onActiveItemChanged = this.onActiveItemChanged.bind(this);
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
      missingField: '',
      items: [],
      columns: [],
      listKey: GuidHelper.getGuid(),
      selection: new Selection()
    };

    this.initItems();
    this.initColumns();
  }

  private initItems() {
    var items = [];
    if (this.state.data != null) {
        this.state.data.map((value: any, index: number) => {
           var item = {};
           this.props.fields.map((field: ICustomListField, indexI: number) => {
            if (value != null && field != null && (field.hidden == null || field.hidden === false)) {
              item[field.title] = value[field.id];
            }
           });
           items.push(item);
        });
    }
    this.state.items = items;
  }

  private initColumns() {
     this.state.columns = buildColumns(this.state.items, true, null, '', false, '', true);
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
      var ctrl = document.getElementById('input-' + this.props.fields[i].id);
      if (ctrl == null)
        continue;
      var str = ctrl['value'];
      if (str.length > 0 && (str[0] == '[' || str[0] == '{'))
        str = JSON.parse(str);

      if (this.props.fields[i].required === true && (str == null || str == '')) {
        this.state.mandatoryOpen = true;
        this.state.missingField = this.props.fields[i].id;
        this.setState(this.state);
        document.getElementById('input-' + this.props.fields[i].id).focus();
        return;
      }

      result[this.props.fields[i].id] = str;
    }
    this.state.data.push(result);
    this.initItems();
    if (this.state.selectedIndex != null && this.state.selectedIndex > 0)
      this.state.selection.setIndexSelected(this.state.selectedIndex, false, false);
    this.state.selectedIndex = null;
    this.setState(this.state);
    this.saveWebPart(this.state.data);

    this.onClickCancel();
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
       this.state.selection.setIndexSelected(this.state.selectedIndex, false, false);
       this.state.selectedIndex = indexToMove - 1;
       this.state.selection.setIndexSelected(this.state.selectedIndex, true, true);
       this.initItems();
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
       this.state.selection.setIndexSelected(this.state.selectedIndex, false, false);
       this.state.selectedIndex = indexToMove + 1;
       this.state.selection.setIndexSelected(this.state.selectedIndex, true, true);
       this.initItems();
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
    this.state.selection.setIndexSelected(this.state.selectedIndex, false, false);
    this.state.data = newData;
    this.state.selectedIndex = null;
    this.initItems();
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
      var ctrl = document.getElementById('input-' + this.props.fields[i].id);
      if (ctrl == null)
        continue;
      var str = ctrl['value'];
      if (str.length > 0 && (str[0] == '[' || str[0] == '{'))
        str = JSON.parse(str);

      if (this.props.fields[i].required === true && (str == null || str == '')) {
        this.state.mandatoryOpen = true;
        this.state.missingField = this.props.fields[i].title;
        this.setState(this.state);
        document.getElementById('input-' + this.props.fields[i].id).focus();
        return;
      }

      result[this.props.fields[i].id] = str;
    }
    this.initItems();
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

  private onActiveItemChanged(item?: any, index?: number, ev?: React.FocusEvent<HTMLElement>): void {
    if (index !== undefined && index >= 0) {
      this.state.selectedIndex = index;
      this.setState(this.state);
    }
    else {
      this.state.selectedIndex = null;
      this.setState(this.state);
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


        <Dialog type={DialogType.close} isOpen={this.state.openPanel} title={this.props.headerText} onDismiss={this.onCancel}
                containerClassName={styles.msDialogMainCustom} isDarkOverlay={true} isBlocking={false}>

          <div style={{width: '630px', height: '500px', overflow: 'scroll'}}>

          { this.state.openListAdd === true ?
          <div>
               {this.props.fields != null ?
              <div>
                <CommandBar
                      isSearchBoxVisible={ false }
                      items={ [
                        { key: 'Add', icon: 'Add', title: strings.CustomListAddItem, name: 'Add', disabled: true, onClick: this.onClickAdd},
                        { key: 'Back', icon: 'Back', title: strings.CustomListBack, name: 'Back', onClick: this.onClickCancel}
                      ] }
                    />
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
                                  <input id={'input-' + value.id} className='ms-TextField-field' style={{marginBottom: '8px'}}/>
                                : ''
                                }
                                { value.type == CustomListFieldType.number ?
                                  <input type="number" role="spinbutton" id={'input-' + value.id} aria-valuemax="99999" aria-valuemin="-999999" aria-valuenow="0" className='ms-TextField-field' style={{width: '100px', marginBottom: '8px'}} />
                                : ''
                                }
                                { value.type == CustomListFieldType.boolean ?
                                  <div  style={{marginBottom: '8px'}}>
                                    <input id={'input-' + value.id}  type="hidden" style={{visibility: 'hidden'}}/>
                                    <input type="radio" role="radio" aria-checked="false" name={'input-' + value.id} style={{width: '18px', height: '18px'}} value={'input-' + value.id} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = true;
                                        }
                                      }
                                    } /> <span style={{fontSize: '14px'}}>{strings.CustomListTrue}</span>
                                    <input type="radio" role="radio"  aria-checked="false" name={'input-' + value.id} style={{width: '18px', height: '18px'}} value={'input-' + value.id} onChange={
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
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDatePickerHost key={'input-' + value.id} label="" properties={this.props.properties} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.dateTime ?
                                  <div>
                                    <input id={'input-' + value.id}  type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDateTimePickerHost key={'input-' + value.id} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.font ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontPickerHost key={'input-' + value.id} label=""  properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.fontSize ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontSizePickerHost key={'input-' + value.id} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.color ?
                                  <div>
                                    <input id={'input-' + value.id} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldColorPickerHost key={'input-' + value.id} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.colorMini ?
                                  <div>
                                    <input id={'input-' + value.id} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldColorPickerMiniHost key={'input-' + value.id} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.icon ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldIconPickerHost key={'input-' + value.id} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.password ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPasswordHost key={'input-' + value.id} label="" properties={this.props.properties}  onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.users ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPeoplePickerHost key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.sharePointGroups ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldGroupPickerHost groupType={IGroupType.SharePoint} key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.securityGroups ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldGroupPickerHost groupType={IGroupType.Security} key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.list ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPListPickerHost key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.folder ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPFolderPickerHost key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.picture ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPicturePickerHost  key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.document ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDocumentPickerHost  key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.officeVideo ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldOfficeVideoPickerHost panelTitle='Select a video'  key={'input-' + value.id} label="" properties={this.props.properties}   context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.stars ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldStarRatingHost  key={'input-' + value.id} label="" properties={this.props.properties} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
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
                <Button style={{marginRight: '10px'}} buttonType={ButtonType.primary} onClick={this.onClickAdd}>{strings.CustomListOK}</Button>
                <Button buttonType={ButtonType.normal} onClick={this.onClickCancel}>{strings.CustomListCancel}</Button>
              </div>
          </div>
          : ''}

          { this.state.editOpen === true ?
          <div>
               {this.props.fields != null ?
                  <div>
                    <CommandBar
                      isSearchBoxVisible={ false }
                      items={ [
                        { key: 'Edit', icon: 'Edit', title: strings.CustomListEdit, name: 'Edit', disabled: true, onClick: this.onClickEdit},
                        { key: 'Back', icon: 'Back', title: strings.CustomListBack, name: 'Back', onClick: this.onClickCancel}
                      ] }
                    />
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
                                  <input id={'input-' + value.id} className='ms-TextField-field' style={{marginBottom: '8px'}} defaultValue={this.state.data[this.state.selectedIndex][value.id]} />
                                : ''
                                }
                                { value.type == CustomListFieldType.number ?
                                  <input type="number" role="spinbutton" id={'input-' + value.id} className='ms-TextField-field' defaultValue={this.state.data[this.state.selectedIndex][value.id]} aria-valuemax="99999" aria-valuemin="-999999" aria-valuenow={this.state.data[this.state.selectedIndex][value.id]} style={{width: '100px', marginBottom: '8px'}} />
                                : ''
                                }
                                { value.type == CustomListFieldType.boolean ?
                                  <div  style={{marginBottom: '8px'}}>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <input type="radio" role="radio" name={'input-' + value.id} style={{width: '18px', height: '18px'}} value={'input-' + value.id} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = true;
                                        }
                                      }
                                    }
                                    defaultChecked={this.state.data[this.state.selectedIndex][value.id] == "true"}
                                    aria-checked={this.state.data[this.state.selectedIndex][value.id] == "true"}
                                    />
                                     <span style={{fontSize: '14px'}}>{strings.CustomListTrue}</span>
                                    <input type="radio" role="radio" name={'input-' + value.id} style={{width: '18px', height: '18px'}} value={'input-' + value.id} onChange={
                                      (elm:any) => {
                                        if (elm.currentTarget.checked == true) {
                                            var name = elm.currentTarget.value;
                                            var input = document.getElementById(name);
                                            input['value'] = false;
                                        }
                                      }
                                    }
                                    defaultChecked={this.state.data[this.state.selectedIndex][value.id] == "false"}
                                    aria-checked={this.state.data[this.state.selectedIndex][value.id] == "false"}
                                    /> <span style={{fontSize: '14px'}}>{strings.CustomListFalse}</span>
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.date ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDatePickerHost key={'input-' + value.id} properties={this.props.properties}  initialDate={this.state.data[this.state.selectedIndex][value.id]} label="" onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.dateTime ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDateTimePickerHost key={'input-' + value.id} properties={this.props.properties}  initialDate={this.state.data[this.state.selectedIndex][value.id]} label="" onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.font ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontPickerHost key={'input-' + value.id} label="" properties={this.props.properties}  initialValue={this.state.data[this.state.selectedIndex][value.id]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.fontSize ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldFontSizePickerHost  key={'input-' + value.id} properties={this.props.properties} label="" initialValue={this.state.data[this.state.selectedIndex][value.id]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.color ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldColorPickerHost key={'input-' + value.id} properties={this.props.properties}  label="" initialColor={this.state.data[this.state.selectedIndex][value.id]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.colorMini ?
                                  <div style={{marginBottom: '5px'}}>
                                    <input id={'input-' + value.id} type="hidden" style={{visibility: 'hidden'}}/>
                                    <PropertyFieldColorPickerMiniHost key={'input-' + value.id} properties={this.props.properties}  label="" initialColor={this.state.data[this.state.selectedIndex][value.id]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.icon ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldIconPickerHost key={'input-' + value.id} properties={this.props.properties}  label="" initialValue={this.state.data[this.state.selectedIndex][value.id]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.password ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPasswordHost key={'input-' + value.id} properties={this.props.properties}  label="" initialValue={this.state.data[this.state.selectedIndex][value.id]} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.users ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={JSON.stringify(this.state.data[this.state.selectedIndex][value.id])}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPeoplePickerHost key={'input-' + value.id} properties={this.props.properties}  label="" initialData={this.state.data[this.state.selectedIndex][value.id]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.securityGroups ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={JSON.stringify(this.state.data[this.state.selectedIndex][value.id])}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldGroupPickerHost groupType={IGroupType.Security} key={'input-' + value.id} properties={this.props.properties}  label="" initialData={this.state.data[this.state.selectedIndex][value.id]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.sharePointGroups ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={JSON.stringify(this.state.data[this.state.selectedIndex][value.id])}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldGroupPickerHost groupType={IGroupType.SharePoint} key={'input-' + value.id} properties={this.props.properties}  label="" initialData={this.state.data[this.state.selectedIndex][value.id]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChangeJson} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.list ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPListPickerHost properties={this.props.properties}  label="" selectedList={this.state.data[this.state.selectedIndex][value.id]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id} key={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.folder ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]}  style={{visibility: 'hidden'}}/>
                                    <PropertyFieldSPFolderPickerHost key={'input-' + value.id} properties={this.props.properties}  label="" initialFolder={this.state.data[this.state.selectedIndex][value.id]}  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.picture ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldPicturePickerHost key={'input-' + value.id} properties={this.props.properties}  label=""  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.document ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldDocumentPickerHost key={'input-' + value.id} properties={this.props.properties}  label=""  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.stars ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldStarRatingHost initialValue={Number(this.state.data[this.state.selectedIndex][value.id])}  key={'input-' + value.id} properties={this.props.properties}  label=""  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
                                  </div>
                                : ''
                                }
                                { value.type == CustomListFieldType.officeVideo ?
                                  <div>
                                    <input id={'input-' + value.id} type="hidden" defaultValue={this.state.data[this.state.selectedIndex][value.id]} style={{visibility: 'hidden'}}/>
                                    <PropertyFieldOfficeVideoPickerHost initialValue={this.state.data[this.state.selectedIndex][value.id]}  panelTitle='Select a video' key={'input-' + value.id} properties={this.props.properties}  label=""  context={this.props.context} onDispose={null} onRender={null} onPropertyChange={this.onPropertyChange} targetProperty={'input-' + value.id}  />
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
                <Button style={{marginRight: '10px'}} buttonType={ButtonType.primary} onClick={this.onClickUpdate}>{strings.CustomListOK}</Button>
                <Button buttonType={ButtonType.normal} onClick={this.onClickCancel}>{strings.CustomListCancel}</Button>
              </div>



          </div>
          : ''}

          { this.state.openListView === true ?
          <div>
              <CommandBar
                isSearchBoxVisible={ false }
                items={ [
                  { key: 'Add', icon: 'Add', title: strings.CustomListAddItem, name: 'Add', onClick: this.onClickAddItem},
                  { key: 'Edit', icon: 'Edit', title: strings.CustomListEdit, name: 'Edit', onClick: this.onClickEdit, disabled: this.state.selectedIndex == null || this.state.selectedIndex < 0 ? true:false},
                  { key: 'Delete', icon: 'Delete', title: strings.CustomListDel, name: 'Delete', onClick: this.onClickDeleteItem, disabled: this.state.selectedIndex == null || this.state.selectedIndex < 0 ? true:false},
                  { key: 'Up', icon: 'ChevronUp', title: '', name: '', onClick: this.onClickMoveUp, disabled: this.state.selectedIndex == null || this.state.selectedIndex <= 0 ? true:false},
                  { key: 'Down', icon: 'ChevronDown', title: '', name: '', onClick: this.onClickMoveDown, disabled: this.state.selectedIndex == null || this.state.selectedIndex < 0 || this.state.selectedIndex >= (this.state.data.length - 1) ? true:false}
                 ] }
              />

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

                  <div style={{marginTop: '20px'}}>

                    <DetailsList
                      setKey={ this.state.listKey }
                      items={ this.state.items }
                      columns={ this.state.columns }
                      checkboxVisibility={ CheckboxVisibility.onHover }
                      layoutMode={ LayoutMode.justified }
                      isHeaderVisible={ true }
                      selection={ this.state.selection }
                      selectionMode={ SelectionMode.single }
                      constrainMode={ ConstrainMode.unconstrained }
                      onActiveItemChanged= { this.onActiveItemChanged }
                      initialFocusedIndex={ this.state.selectedIndex }
                    />

                  </div>
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
