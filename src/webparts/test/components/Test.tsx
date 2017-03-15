import * as React from 'react';

import styles from '../Test.module.scss';
import { ITestWebPartProps } from '../ITestWebPartProps';
import { IPropertyFieldPeople } from '../../../PropertyFieldPeoplePicker';

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

              <div className="ms-MessageBar">
                <div className="ms-MessageBar-content">
                  <div className="ms-MessageBar-icon">
                    <i className="ms-Icon ms-Icon--Info"></i>
                  </div>
                  <div className="ms-MessageBar-text">
                    Edit this WebPart to test the custom fields.
                  </div>
                </div>
              </div>

              <p className="ms-fontSize-xxl">
                <i className="ms-Icon ms-Icon--ClearFormatting" aria-hidden="true"></i>
                    &nbsp; Layout Fields
              </p>

              <p>
                <i className="ms-Icon ms-Icon--Font" aria-hidden="true"></i>&nbsp;
                <b>Font</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldFontPicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.font}

              </p>
              <p>
                <i className="ms-Icon ms-Icon--FontSize" aria-hidden="true"></i>&nbsp;
                <b>Font Size</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldFontSizePicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.fontSize}

              </p>
              <p >
                <i className="ms-Icon ms-Icon--Color" aria-hidden="true"></i>&nbsp;
                <b>Color</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldColorPicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.color}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--Waffle" aria-hidden="true"></i>&nbsp;
                <b>Icon</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldIconPicker" target="_doc">(Doc)</a> : &nbsp;
                <i className={'ms-Icon ' + this.props.icon} aria-hidden="true" style={{fontSize:'large'}}></i>
                &nbsp;{this.props.icon}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--AlignLeft" aria-hidden="true"></i>&nbsp;
                <b>Align</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldAlignPicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.align}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--FullScreen" aria-hidden="true"></i>&nbsp;
                <b>Dimension</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDimensionPicker" target="_doc">(Doc)</a> : &nbsp;
                {this.props.dimension != null ? this.props.dimension.width + ' x ' + this.props.dimension.height : ''}
              </p>

              <p className="ms-fontSize-xxl">
               <i className="ms-Icon ms-Icon--InsertTextBox" aria-hidden="true"></i>
                &nbsp; Text Input Fields
              </p>
              <div>
                <i className="ms-Icon ms-Icon--ChevronDown" aria-hidden="true"></i>&nbsp;
                <b>DropDown Select</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDropDownSelect" target="_doc">(Doc)</a> : &nbsp;
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
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldSortableList" target="_doc">(Doc)</a> :&nbsp;
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
                <b>Treeview</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldTreeView" target="_doc">(Doc)</a> :&nbsp;
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
              <p >
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Custom List</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldCustomList" target="_doc">(Doc)</a> : &nbsp;
                {JSON.stringify(this.props.customList)}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--Lock" aria-hidden="true"></i>&nbsp;
                <b>Password</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldPassword" target="_doc">(Doc)</a> : &nbsp;
                {this.props.password}
              </p>
               <p >
                <i className="ms-Icon ms-Icon--Font" aria-hidden="true"></i>&nbsp;
                <b>Rich Text Box</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldRichTextBox" target="_doc">(Doc)</a> : &nbsp;
                {this.props.richTextBox}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Date</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDatePicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.date}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Date ISO</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDatePicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.date2}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Date Time</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDateTimePicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.datetime}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--CalendarWorkWeek" aria-hidden="true"></i>&nbsp;
                <b>Slider Range</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldSliderRange" target="_doc">(Doc)</a> :&nbsp;
                {this.props.sliderRange}
              </p>

              <p>
                <i className="ms-Icon ms-Icon--Phone" aria-hidden="true"></i>&nbsp;
                <b>Phone Number</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldPhoneNumber" target="_doc">(Doc)</a> :&nbsp;
                {this.props.phone}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--PaymentCard" aria-hidden="true"></i>&nbsp;
                <b>Credit Card</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldMaskedInput" target="_doc">(Doc)</a> :&nbsp;
                {this.props.maskedInput}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--MapPin" aria-hidden="true"></i>&nbsp;
                <b>Geolocation</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldMapPicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.geolocation}
              </p>

              <p className="ms-fontSize-xxl">
               <i className="ms-Icon ms-Icon--DeveloperTools" aria-hidden="true"></i>
                 &nbsp; SharePoint Fields
              </p>
              <div>
                <i className="ms-Icon ms-Icon--PeopleAdd" aria-hidden="true"></i>&nbsp;
                <b>Users</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldPeoplePicker" target="_doc">(Doc)</a> :&nbsp;

                <ul>
                {
                  peopleList.map((element: IPropertyFieldPeople, i:number) => {
                    return (
                      <li>
                        Username : {element.fullName}<br/>
                        Login: {element.login}<br/>
                        Email: {element.email}<br/>
                        JobTitle: {element.jobTitle}<br/>
                      </li>
                    );
                })}
                </ul>
              </div>

              <p>
                <i className="ms-Icon ms-Icon--Picture" aria-hidden="true"></i>&nbsp;
                <b>Picture</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldPicturePicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.picture}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--Document" aria-hidden="true"></i>&nbsp;
                <b>Document</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDocumentPicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.document}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>List</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldSPListPicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.list}
              </p>
              <p>
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Query</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldSPListQuery" target="_doc">(Doc)</a> :&nbsp;
                {this.props.query}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Display Mode</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldDisplayMode" target="_doc">(Doc)</a> : &nbsp;
                {this.props.displayMode}
              </p>
              <p >
                <i className="ms-Icon ms-Icon--FolderSearch" aria-hidden="true"></i>&nbsp;
                <b>Folder</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldSPFolderPicker" target="_doc">(Doc)</a> :&nbsp;
                {this.props.folder}
              </p>
              <div >
                <i className="ms-Icon ms-Icon--List" aria-hidden="true"></i>&nbsp;
                <b>Lists</b>
                &nbsp;<a className="ms-fontSize-sPlus" href="https://github.com/OlivierCC/sp-client-custom-fields/wiki/PropertyFieldSPListMultiplePicker" target="_doc">(Doc)</a> :&nbsp;
                <ul>
                {
                  lists.map((element: string, i:number) => {
                    return (
                      <li>{element}</li>
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




