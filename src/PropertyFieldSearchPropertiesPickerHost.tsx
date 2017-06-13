/**
 * @file PropertyFieldSearchPropertiesPickerHost.tsx
 * Renders the controls for PropertyFieldSearchPropertiesPicker component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldSearchPropertiesPickerPropsInternal } from './PropertyFieldSearchPropertiesPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { TagPicker, IBasePickerSuggestionsProps, ITag } from 'office-ui-fabric-react/lib/Pickers';

/**
 * @interface
 * PropertyFieldSearchPropertiesPickerHost properties interface
 *
 */
export interface IPropertyFieldSearchPropertiesPickerHostProps extends IPropertyFieldSearchPropertiesPickerPropsInternal {
}

export interface IPropertyFieldSearchPropertiesPickerState {
  errorMessage: string;
  properties: ITag[];
  selectedProperties: ITag[];
}

/**
 * @class
 * Renders the controls for PropertyFieldSearchPropertiesPicker component
 */
export default class PropertyFieldSearchPropertiesPickerHost extends React.Component<IPropertyFieldSearchPropertiesPickerHostProps, IPropertyFieldSearchPropertiesPickerState> {

  private defaultProperties: string[] = [
    "AboutMe","Account","AccountName","acronymaggre","AssignedTo","AttachmentType","Author","BaseOfficeLocation","CategoryNavigationUrl","charset","Colleagues","CombinedUserProfileNames","companies","contentclass","ContentsHidden","ContentSource","ContentType","ContentTypeId","Created","CreatedBy","Date00","Date01","Date02","Date03","Date04","Date05","Date06","Date07","Date08","Date09","Decimal00","Decimal01","Decimal02","Decimal03","Decimal04","Decimal05","Decimal06","Decimal07","Decimal08","Decimal09","deeplinks","defaggre","Department","Description","DetectedLanguage","DisplayAuthor","DisplayDate","DMSDocAccessRight","DMSDocAuthor","DMSDocTitle","docacl","DocComments","DocId","DocKeywords","DocSignature","DocSubject","DocumentSignature","domain","Double00","Double01","Double02","Double03","Double04","Double05","Double06","Double07","Double08","Double09","DuplicateHash","EduAssignmentCategory","EduAssignmentFormat","EduMaximumScore","EndDate","ExpirationTime","ExtractedAuthor","ExtractedDate","FileExtension","Filename","FileType","FirstLevelColleagues","FirstLevelMutualFollowings","FirstName","FollowAllAnchor","format","GeneratedTitle","Genre","HierarchyUrl","HitHighlightedProperties","HitHighlightedSummary","HostingPartition","hwboost","ImageDateCreated","importance","Int00","Int01","Int02","Int03","Int04","Int05","Int06","Int07","Int08","Int09","Int10","Int11","Int12","Int13","Int14","Int15","Int16","Int17","Int18","Int19","Int20","Int21","Int22","Int23","Int24","Int25","Int26","Int27","Int28","Int29","Int30","Int31","Int32","Int33","Int34","Int35","Int36","Int37","Int38","Int39","Int40","Int41","Int42","Int43","Int44","Int45","Int46","Int47","Int48","Int49","Interests","IsContainer","IsData","IsDocument","IsMyDocuments","IsPublishingCatalog","IsReport","JobTitle","Keywords","language","languages","LastModifiedTime","LastName","ListID","ListItemID","ListUrl","Location","ManagedProperties","MediaDuration","Memberships","MetadataAuthor","MicroBlogType","MobilePhone","ModifiedBy","NLCodePage","Notes","OfficeNumber","OrgNames","OrgParentNames","OrgParentUrls","OrgUrls","OWS_URL","owsmetadatafacetinfo","owstaxidmetadataalltagsinfo","owstaxIdProductCatalogItemCategory","ParentLink","PastProjects","Path","People","PeopleInMedia","PeopleKeywords","PhoneNumber","PictureHeight","PictureThumbnailURL","PictureURL","PictureWidth","PostAuthor","PreferredName","Priority","PrivateColleagues","processingtime","ProductCatalogGroupNumberOWSTEXT","ProfileExpertise","ProfileName","Pronunciations","Purpose","RankDetail","RankingWeightHigh","RankingWeightLow","RankingWeightName","recommendedfor","RefinableDate00","RefinableDate01","RefinableDate02","RefinableDate03","RefinableDate04","RefinableDate05","RefinableDate06","RefinableDate07","RefinableDate08","RefinableDate09","RefinableDate10","RefinableDate11","RefinableDate12","RefinableDate13","RefinableDate14","RefinableDate15","RefinableDate16","RefinableDate17","RefinableDate18","RefinableDate19","RefinableDecimal00","RefinableDecimal01","RefinableDecimal02","RefinableDecimal03","RefinableDecimal04","RefinableDecimal05","RefinableDecimal06","RefinableDecimal07","RefinableDecimal08","RefinableDecimal09","RefinableDouble00","RefinableDouble01","RefinableDouble02","RefinableDouble03","RefinableDouble04","RefinableDouble05","RefinableDouble06","RefinableDouble07","RefinableDouble08","RefinableDouble09","RefinableInt00","RefinableInt01","RefinableInt02","RefinableInt03","RefinableInt04","RefinableInt05","RefinableInt06","RefinableInt07","RefinableInt08","RefinableInt09","RefinableInt10","RefinableInt11","RefinableInt12","RefinableInt13","RefinableInt14","RefinableInt15","RefinableInt16","RefinableInt17","RefinableInt18","RefinableInt19","RefinableInt20","RefinableInt21","RefinableInt22","RefinableInt23","RefinableInt24","RefinableInt25","RefinableInt26","RefinableInt27","RefinableInt28","RefinableInt29","RefinableInt30","RefinableInt31","RefinableInt32","RefinableInt33","RefinableInt34","RefinableInt35","RefinableInt36","RefinableInt37","RefinableInt38","RefinableInt39","RefinableInt40","RefinableInt41","RefinableInt42","RefinableInt43","RefinableInt44","RefinableInt45","RefinableInt46","RefinableInt47","RefinableInt48","RefinableInt49","RefinableString00","RefinableString01","RefinableString02","RefinableString03","RefinableString04","RefinableString05","RefinableString06","RefinableString07","RefinableString08","RefinableString09","RefinableString10","RefinableString11","RefinableString12","RefinableString13","RefinableString14","RefinableString15","RefinableString16","RefinableString17","RefinableString18","RefinableString19","RefinableString20","RefinableString21","RefinableString22","RefinableString23","RefinableString24","RefinableString25","RefinableString26","RefinableString27","RefinableString28","RefinableString29","RefinableString30","RefinableString31","RefinableString32","RefinableString33","RefinableString34","RefinableString35","RefinableString36","RefinableString37","RefinableString38","RefinableString39","RefinableString40","RefinableString41","RefinableString42","RefinableString43","RefinableString44","RefinableString45","RefinableString46","RefinableString47","RefinableString48","RefinableString49","RefinableString50","RefinableString51","RefinableString52","RefinableString53","RefinableString54","RefinableString55","RefinableString56","RefinableString57","RefinableString58","RefinableString59","RefinableString60","RefinableString61","RefinableString62","RefinableString63","RefinableString64","RefinableString65","RefinableString66","RefinableString67","RefinableString68","RefinableString69","RefinableString70","RefinableString71","RefinableString72","RefinableString73","RefinableString74","RefinableString75","RefinableString76","RefinableString77","RefinableString78","RefinableString79","RefinableString80","RefinableString81","RefinableString82","RefinableString83","RefinableString84","RefinableString85","RefinableString86","RefinableString87","RefinableString88","RefinableString89","RefinableString90","RefinableString91","RefinableString92","RefinableString93","RefinableString94","RefinableString95","RefinableString96","RefinableString97","RefinableString98","RefinableString99","Responsibilities","RobotsNoIndex","RootPostID","RootPostOwnerID","RootPostUniqueID","Schools","SecondaryFileExtension","SecondLevelColleagues","ServerRedirectedURL","ServiceApplicationID","SharedWithInternal","SipAddress","Site","SiteClosed","SiteID","sitename","SiteTitle","Size","Skills","SocialTag","SocialTagTextUrl","SPContentType","SPSiteURL","StartDate","Status","Tags","Title","tld","UrlDepth","urlkeywords","urls","UsageAnalyticsId","UsageEventItemId","UserName","UserProfile_GUID","WebId","WebTemplate","WikiCategory","WordCustomRefiner1","WordCustomRefiner2","WordCustomRefiner3","WordCustomRefiner4","WordCustomRefiner5","WordExactCustomRefiner","WordPartCustomRefiner1","WordPartCustomRefiner2","WordPartCustomRefiner3","WordPartCustomRefiner4","WordPartCustomRefiner5","WordPartExactCustomRefiner","WorkEmail","WorkPhone","YomiDisplayName"
  ];


  private async: Async;
  private delayedValidate: (value: string[]) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldSearchPropertiesPickerHostProps) {
    super(props);

    this.async = new Async(this);
    this.state = {
      errorMessage: '',
      properties: this.getDefaultProperties(),
      selectedProperties: this.getSelectedProperties()
    };

    //Bind the current object to the external called onSelectDate method
    this.onItemChanged = this.onItemChanged.bind(this);
    this.onFilterChanged = this.onFilterChanged.bind(this);
    this.listContainsTag = this.listContainsTag.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  private getDefaultProperties(): ITag[] {
    var res: ITag[] = [];
    for (var i = 0; i < this.defaultProperties.length; i++) {
      var tag: ITag = {
        key: this.defaultProperties[i],
        name: this.defaultProperties[i]
      };
      res.push(tag);
    }
    return res;
  }

  private getSelectedProperties(): ITag[] {
    var res: ITag[] = [];
    if (this.props.selectedProperties === undefined)
      return res;
    for (var i = 0; i < this.props.selectedProperties.length; i++) {
      var tag: ITag = {
        key: this.props.selectedProperties[i],
        name: this.props.selectedProperties[i]
      };
      res.push(tag);
    }
    return res;
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedProperties, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedProperties, value);
        this.setState({ errorMessage: result} as IPropertyFieldSearchPropertiesPickerState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedProperties, value);
          this.setState({ errorMessage } as IPropertyFieldSearchPropertiesPickerState);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedProperties, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string[], newValue: string[]) {
    this.props.properties[this.props.targetProperty] = newValue;
    this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    if (!this.props.disableReactivePropertyChanges && this.props.render != null)
        this.props.render();
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    if (this.async !== undefined)
      this.async.dispose();
  }

  /**
   * @function
   * Called when the TagPicker text changed
   * @param filterText
   * @param tagList
   */
  private onFilterChanged(filterText: string, tagList: ITag[]) {
    return filterText ? this.state.properties.filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0).filter(item => !this.listContainsTag(item, tagList)) : [];
  }

  /**
   * @function
   * Tests if the selected list contains already the tag
   * @param tag
   * @param tagList
   */
  private listContainsTag(tag: ITag, tagList: ITag[]) {
    if (!tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
  }

  /**
   * @function
   * Occurs when the list of selected items changed
   * @param selectedItems
   */
  private onItemChanged(selectedItems: ITag[]): void {
    var res: string[] = [];
    for (var i = 0; i < selectedItems.length; i++) {
      res.push(selectedItems[i].key);
    }
    this.delayedValidate(res);
  }

  /**
   * @function
   * Renders the controls
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <TagPicker
          onResolveSuggestions={ this.onFilterChanged }
          getTextFromItem= {(item: ITag) => { return item.name; } }
          defaultSelectedItems={this.state.selectedProperties}
          onChange={this.onItemChanged}
          pickerSuggestionsProps={
            {
              suggestionsHeaderText: this.props.suggestionsHeaderText,
              noResultsFoundText: this.props.noResultsFoundText,
              loadingText: this.props.loadingText
            }
          }
        />
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
          <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
          <span>
            <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
          </span>
          </div>
        : ''}
      </div>
    );
  }
}
