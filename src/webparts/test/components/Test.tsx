import * as React from 'react';

import styles from '../Test.module.scss';
import { MessageBar } from 'office-ui-fabric-react/lib/MessageBar';

import { ITestWebPartProps } from '../ITestWebPartProps';
import { IPropertyFieldPeople } from '../../../PropertyFieldPeoplePicker';
import { IPropertyFieldGroup } from '../../../PropertyFieldGroupPicker';
import { ISPTermSet } from '../../../PropertyFieldTermSetPicker';

export interface ITestProps extends ITestWebPartProps {
}

export default class Test extends React.Component<ITestProps, {}> {
  public render(): JSX.Element {

    var peopleList: IPropertyFieldPeople[] = [];
    if (this.props && this.props.people)
      peopleList = this.props.people;
    var lists: string[] = [];
    if (this.props && this.props.listsCollection)
      lists = this.props.listsCollection;

    return (
      <div className={styles.test}>

        <div className={styles.container}>
          <div>
            <div style={{ backgroundColor: this.props.color, fontFamily: this.props.font, fontSize: this.props.fontSize ? this.props.fontSize : '12px', padding: '20px' }}>

            <MessageBar>
               Edit this WebPart to test the custom fields.
            </MessageBar>

              <p className="ms-fontSize-xxl">
                <i className="ms-Icon ms-Icon--ClearFormatting" aria-hidden="true"></i>
                    &nbsp; Layout Fields
              </p>

              <p>
                <i className="ms-Icon ms-Icon--Font" aria-hidden="true"></i>&nbsp;
                <b>Font</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldfontpicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.font}

              </p>
              <p>
                <i className="ms-Icon ms-Icon--FontSize" aria-hidden="true"></i>&nbsp;
                <b>Font Size</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldfontsizepicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.fontSize}

              </p>
              <p >
                <i className="ms-Icon ms-Icon--Color" aria-hidden="true"></i>&nbsp;
                <b>Color (Mini)</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldcolorpickermini" target="_doc">(Doc)</a> : &nbsp;
                {this.props.miniColor}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--Color" aria-hidden="true"></i>&nbsp;
                <b>Color</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldcolorpicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.color}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--Waffle" aria-hidden="true"></i>&nbsp;
                <b>Icon</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldiconpicker" target="_doc">(Doc)</a> : &nbsp;
                <i className={'ms-Icon ' + this.props.icon} aria-hidden="true" style={{fontSize:'large'}}></i>
                &nbsp;{this.props.icon}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--AlignLeft" aria-hidden="true"></i>&nbsp;
                <b>Align</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldalignpicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.align}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--FullScreen" aria-hidden="true"></i>&nbsp;
                <b>Dimension</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddimensionpicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.dimension != null ? this.props.dimension.width + ' x ' + this.props.dimension.height : ''}
              </p>

              <p className="ms-fontSize-xxl">
               <i className="ms-Icon ms-Icon--InsertTextBox" aria-hidden="true"></i>
                &nbsp; Text Input Fields
              </p>
              <div>
                <i className="ms-Icon ms-Icon--ChevronDown" aria-hidden="true"></i>&nbsp;
                <b>DropDown Select</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddropdownselect" target="_doc">(Doc)</a> : &nbsp;
                <ul>
                { this.props.dropDownSelect != null ?
                  this.props.dropDownSelect.map((element: string, i:number) => {
                    return (
                      <li key={i}>{element}</li>
                    );
                  })
                  : ''
                }
                </ul>

              </div>
              <div>
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Sortable List</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldsortablelist" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                { this.props.sortableList != null ?
                  this.props.sortableList.map((element: string, i:number) => {
                    return (
                      <li key={i}>{element}</li>
                    );
                  })
                  : ''
                }
                </ul>
              </div>
              <div>
                <i className="ms-Icon ms-Icon--Breadcrumb" aria-hidden="true"></i>&nbsp;
                <b>DropDown Treeview</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddropdowntreeview" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                { this.props.dropDownTreeView != null ?
                  this.props.dropDownTreeView.map((element: string, i:number) => {
                    return (
                      <li key={i}>{element}</li>
                    );
                  })
                  : ''
                }
                </ul>
              </div>
              <div>
                <i className="ms-Icon ms-Icon--Breadcrumb" aria-hidden="true"></i>&nbsp;
                <b>Treeview</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldtreeview" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                { this.props.treeView != null ?
                  this.props.treeView.map((element: string, i:number) => {
                    return (
                      <li key={i}>{element}</li>
                    );
                  })
                  : ''
                }
                </ul>
              </div>
              <div>
                <i className="ms-Icon ms-Icon--Tag" aria-hidden="true"></i>&nbsp;
                <b>Tags</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldtagpicker" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                { this.props.tags != null ?
                  this.props.tags.map((element: any, i:number) => {
                    return (
                      <li key={i}>{element.name}</li>
                    );
                  })
                  : ''
                }
                </ul>
              </div>
              <p >
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Custom List</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldcustomlist" target="_doc">(Doc)</a> : &nbsp;
                {JSON.stringify(this.props.customList)}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--FavoriteStarFill" aria-hidden="true"></i>&nbsp;
                <b>Star Rating</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldstarrating" target="_doc">(Doc)</a> : &nbsp;
                {this.props.starRating}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--Lock" aria-hidden="true"></i>&nbsp;
                <b>Password</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldpassword" target="_doc">(Doc)</a> : &nbsp;
                {this.props.password}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--NumberField" aria-hidden="true"></i>&nbsp;
                <b>Auto Complete Text</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldautocomplete" target="_doc">(Doc)</a> : &nbsp;
                {this.props.autoSuggest}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--NumberField" aria-hidden="true"></i>&nbsp;
                <b>Number</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldnumericinput" target="_doc">(Doc)</a> : &nbsp;
                {this.props.numeric}
              </p>
               <p >
                <i className="ms-Icon ms-Icon--Font" aria-hidden="true"></i>&nbsp;
                <b>Rich Text Box</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldrichtextbox" target="_doc">(Doc)</a> : &nbsp;
                {this.props.richTextBox}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Date</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddatepicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.date}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Date ISO</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddatepicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.date2}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Date Time</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddatetimepicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.datetime}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Slider Range</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldsliderrange" target="_doc">(Doc)</a> :&nbsp;
                {this.props.sliderRange}
              </p>

              <p>
                <i className="ms-Icon ms-Icon--Phone" aria-hidden="true"></i>&nbsp;
                <b>Phone Number</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldphonenumber" target="_doc">(Doc)</a> :&nbsp;
                {this.props.phone}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--PaymentCard" aria-hidden="true"></i>&nbsp;
                <b>Credit Card</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldmaskedinput" target="_doc">(Doc)</a> :&nbsp;
                {this.props.maskedInput}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--MapPin" aria-hidden="true"></i>&nbsp;
                <b>Geolocation</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldmappicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.geolocation}
              </p>

              <p className="ms-fontSize-xxl">
               <i className="ms-Icon ms-Icon--DeveloperTools" aria-hidden="true"></i>
                 &nbsp; SharePoint Fields
              </p>
              <div>
                <i className="ms-Icon ms-Icon--PeopleAdd" aria-hidden="true"></i>&nbsp;
                <b>Users</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldpeoplepicker" target="_doc">(Doc)</a> :&nbsp;

                <ul>
                {
                  peopleList.map((element: IPropertyFieldPeople, i:number) => {
                    return (
                      <li key={'people' + i}>
                        Username : {element.fullName}<br/>
                        Login: {element.login}<br/>
                        Email: {element.email}<br/>
                        JobTitle: {element.jobTitle}<br/>
                      </li>
                    );
                })}
                </ul>
              </div>
              <div>
                <i className="ms-Icon ms-Icon--PeopleAdd" aria-hidden="true"></i>&nbsp;
                <b>Groups</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldgrouppicker" target="_doc">(Doc)</a> :&nbsp;

                <ul>
                {
                  this.props.groups !== undefined ? this.props.groups.map((element: IPropertyFieldGroup, i:number) => {
                    return (
                      <li key={'groups' + i}>
                        FullName : {element.fullName}<br/>
                        Description: {element.description}<br/>
                        Login: {element.login}<br/>
                        ID: {element.id}<br/>
                      </li>
                    );
                }) : ''}
                </ul>
              </div>

              <p>
                <i className="ms-Icon ms-Icon--Picture" aria-hidden="true"></i>&nbsp;
                <b>Picture</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldpicturepicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.picture}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--Document" aria-hidden="true"></i>&nbsp;
                <b>Document</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddocumentepicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.document}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--VideoSolid" aria-hidden="true"></i>&nbsp;
                <b>Office Video</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldofficevideopicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.officeVideo}
              </p>
              <div>
                <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>&nbsp;
                <b>Search Properties</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldsearchpropertiespicker" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                {
                  this.props.searchProperties !== undefined ? this.props.searchProperties.map((element: string, i:number) => {
                    return (
                      <li key={'groups' + i}>{element}<br/>
                      </li>
                    );
                }) : ''}
                </ul>
              </div>
              <div>
                <i className="ms-Icon ms-Icon--Tag" aria-hidden="true"></i>&nbsp;
                <b>Term Sets</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldtermsetpicker" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                {
                  this.props.termSets !== undefined ? this.props.termSets.map((element: ISPTermSet, i:number) => {
                    return (
                      <li key={'groups' + i}>Name: {element.Name}, Guid: {element.Guid}, TermStore Guid: {element.TermStoreGuid}<br/>
                      </li>
                    );
                }) : ''}
                </ul>
              </div>
              <p>
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>List</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldsplistpicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.list}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Query</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldsplistquery" target="_doc">(Doc)</a> :&nbsp;
                {this.props.query}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Display Mode</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfielddisplaymode" target="_doc">(Doc)</a> : &nbsp;
                {this.props.displayMode}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--FolderSearch" aria-hidden="true"></i>&nbsp;
                <b>Folder</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldspfolderpicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.folder}
              </p>
              <div >
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Lists</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://oliviercc.github.io/sp-client-custom-fields/propertyfieldsplistmultiplepicker" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                {
                  lists.map((element: string, i:number) => {
                    return (
                      <li key={'list' + i}>{element}</li>
                    );
                })}
                </ul>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}




