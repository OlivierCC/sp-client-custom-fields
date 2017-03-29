/**
 * @file PropertyFieldDropDownTreeViewHost.tsx
 * Renders the controls for PropertyFieldDropDownTreeView component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDropDownTreeViewPropsInternal, IDropDownTreeViewNode } from './PropertyFieldDropDownTreeView';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import GuidHelper from './GuidHelper';

require('react-ui-tree-draggable/dist/react-ui-tree.css');
var Tree: any = require('react-ui-tree-draggable/dist/react-ui-tree');

/**
 * @interface
 * PropertyFieldDropDownTreeViewHost properties interface
 *
 */
export interface IPropertyFieldDropDownTreeViewHostProps extends IPropertyFieldDropDownTreeViewPropsInternal {
}

/**
 * @interface
 * PropertyFieldDropDownTreeViewHost state interface
 *
 */
export interface IPropertyFieldDropDownTreeViewHostState {
  isOpen: boolean;
  isHoverDropdown?: boolean;
  errorMessage?: string;
  tree: IDropDownTreeViewNode[];
  activeNodes: IDropDownTreeViewNode[];
}

/**
 * @class
 * Renders the controls for PropertyFieldDropDownTreeView component
 */
export default class PropertyFieldDropDownTreeViewHost extends React.Component<IPropertyFieldDropDownTreeViewHostProps, IPropertyFieldDropDownTreeViewHostState> {

  private async: Async;
  private delayedValidate: (value: string[]) => void;
  private _key: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldDropDownTreeViewHostProps) {
    super(props);

    //Bind the current object to the external called onSelectDate method
    this.onOpenDialog = this.onOpenDialog.bind(this);
    this.mouseEnterDropDown = this.mouseEnterDropDown.bind(this);
    this.mouseLeaveDropDown = this.mouseLeaveDropDown.bind(this);
    this._key = GuidHelper.getGuid();

    //Init the state
    this.state = {
        isOpen: false,
        isHoverDropdown: false,
        errorMessage: '',
        tree: this.props.tree,
        activeNodes: this.getDefaultActiveNodesFromTree()
      };

    this.renderNode = this.renderNode.bind(this);
    this.onClickNode = this.onClickNode.bind(this);
    this.saveSelectedNodes = this.saveSelectedNodes.bind(this);
    this.handleTreeChange = this.handleTreeChange.bind(this);

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  private getDefaultActiveNodesFromTree(): IDropDownTreeViewNode[] {
    var res: IDropDownTreeViewNode[] = [];
    this.props.tree.map((node: IDropDownTreeViewNode) => {
      var subTreeViewNodes: IDropDownTreeViewNode[] = this.getDefaultActiveNodes(node);
      subTreeViewNodes.map((subNode: IDropDownTreeViewNode) => {
        res.push(subNode);
      });
    });
    return res;
  }

  /**
   * @function
   * Gets the list of activated nodes from the  selectedNodesIDs property
   * @param node
   */
  private getDefaultActiveNodes(node: IDropDownTreeViewNode): IDropDownTreeViewNode[] {
    var res: IDropDownTreeViewNode[] = [];
    if (node === undefined || node == null || this.props.selectedNodesIDs === undefined || this.props.selectedNodesIDs == null)
      return res;
    if (this.props.selectedNodesIDs.indexOf(node.id) != -1)
      res.push(node);
    if (node.children !== undefined) {
      for (var i = 0; i < node.children.length; i++) {
        var subTreeViewNodes: IDropDownTreeViewNode[] = this.getDefaultActiveNodes(node.children[i]);
        subTreeViewNodes.map((subNode: IDropDownTreeViewNode) => {
          res.push(subNode);
        });
      }
    }
    return res;
  }

  /**
   * @function
   * Gets the given node position in the active nodes collection
   * @param node
   */
  private getSelectedNodePosition(node: IDropDownTreeViewNode): number {
    for (var i = 0; i < this.state.activeNodes.length; i++) {
      if (node === this.state.activeNodes[i])
        return i;
    }
    return -1;
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedNodesIDs, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedNodesIDs, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedNodesIDs, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedNodesIDs, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string[], newValue: string[]) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    }
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
   * Function to open the dialog
   */
  private onOpenDialog(): void {
    if (this.props.disabled === true)
      return;
    this.state.isOpen = !this.state.isOpen;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is hover the fontpicker
   */
  private mouseEnterDropDown(element?: any) {
    this.state.isHoverDropdown = true;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is leaving the fontpicker
   */
  private mouseLeaveDropDown(element?: any) {
    this.state.isHoverDropdown = false;
    this.setState(this.state);
  }

  /**
   * clicks on a node
   * @param node
   */
  private onClickNode(node: IDropDownTreeViewNode): void {
    if (this.props.allowFoldersSelections === false && (node.children !== undefined && node.children.length != 0))
      return;
    if (this.props.allowMultipleSelections === false) {
      this.state.activeNodes = [node];
    }
    else {
      var index = this.getSelectedNodePosition(node);
      if (index != -1)
        this.state.activeNodes.splice(index, 1);
      else
        this.state.activeNodes.push(node);
    }
    this.setState(this.state);
    this.saveSelectedNodes();
  }

  /**
   * Saves the selected nodes
   */
  private saveSelectedNodes(): void {
    var res: string[] = [];
    for (var i = 0; i < this.state.activeNodes.length; i++) {
      res.push(this.state.activeNodes[i].id);
    }
    this.delayedValidate(res);
  }

  /**
   * @function
   * Renders the given node
   * @param node
   */
  private renderNode(node: IDropDownTreeViewNode): JSX.Element {
    var style: any = { padding: '4px 5px', width: '100%', display: 'block'};
    var selected: boolean = this.getSelectedNodePosition(node) != -1;
    if (selected === true) {
      style.backgroundColor = '#EAEAEA';
    }
    var isFolder: boolean = false;
    if (node.leaf === false || (node.children !== undefined && node.children.length != 0))
      isFolder = true;
    var checkBoxAvailable: boolean = this.props.checkboxEnabled;
    if (this.props.allowFoldersSelections === false && isFolder === true)
      checkBoxAvailable = false;
    var picUrl: string = '';
    if (selected === true && node.selectedPictureUrl !== undefined)
      picUrl = node.selectedPictureUrl;
    else if (node.collapsed !== true && node.expandedPictureUrl !== undefined)
      picUrl = node.expandedPictureUrl;
    else if (node.pictureUrl !== undefined)
      picUrl = node.pictureUrl;
    var nodeStyle = 'ms-Checkbox-label';
    if (selected === true)
      nodeStyle += ' is-checked';
    return (
        <span style={style} onClick={this.onClickNode.bind(null, node)} role="menuitem">
          { checkBoxAvailable ?
              <label className={nodeStyle} style={{padding: 0, margin: 0}} htmlFor={node.id}>
                <input disabled={this.props.disabled} style={{width: '18px', height: '18px', opacity: 0}}
                  checked={selected} aria-checked={selected} readOnly={true}
                  type="checkbox" role="checkbox" />
              </label>
            : ''
          }
          {
            picUrl !== undefined && picUrl != '' ?
              <img src={picUrl} width="18" height="18" style={{paddingRight: '5px'}} alt={node.label}/>
            : ''
          }
          {node.label}
        </span>
    );
  }

  /**
   * Handles tree changes
   * @param rootNode
   * @param index
   */
  private handleTreeChange(rootNode: any, index: number): void {
    this.state.tree[index] = rootNode;
    this.setState(this.state);
  }

  /**
   * @function
   * Renders the control
   */
  public render(): JSX.Element {

      //User wants to use the preview font picker, so just build it
      var fontSelect = {
        fontSize: '16px',
        width: '100%',
        position: 'relative',
        display: 'inline-block',
        zoom: 1
      };
      var dropdownColor = '1px solid #c8c8c8';
      if (this.props.disabled === true)
        dropdownColor = '1px solid #f4f4f4';
      else if (this.state.isOpen === true)
        dropdownColor = '1px solid #3091DE';
      else if (this.state.isHoverDropdown === true)
        dropdownColor = '1px solid #767676';
      var fontSelectA = {
        backgroundColor: this.props.disabled === true ? '#f4f4f4' : '#fff',
        borderRadius        : '0px',
        backgroundClip        : 'padding-box',
        border: dropdownColor,
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        position: 'relative',
        height: '26px',
        lineHeight: '26px',
        padding: '0 0 0 8px',
        color: this.props.disabled === true ? '#a6a6a6' : '#444',
        textDecoration: 'none',
        cursor: this.props.disabled === true ? 'default' : 'pointer'
      };
      var fontSelectASpan = {
        marginRight: '26px',
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        lineHeight: '1.8',
        textOverflow: 'ellipsis',
        cursor: this.props.disabled === true ? 'default' : 'pointer',
        fontWeight: 400
      };
      var fontSelectADiv = {
        borderRadius        : '0 0px 0px 0',
        backgroundClip        : 'padding-box',
        border: '0px',
        position: 'absolute',
        right: '0',
        top: '0',
        display: 'block',
        height: '100%',
        width: '22px'
      };
      var fontSelectADivB = {
        display: 'block',
        width: '100%',
        height: '100%',
        cursor: this.props.disabled === true ? 'default' : 'pointer',
        marginTop: '2px'
      };
      var fsDrop = {
        background: '#fff',
        border: '1px solid #aaa',
        borderTop: '0',
        position: 'absolute',
        top: '29px',
        left: '0',
        width: 'calc(100% - 2px)',
        //boxShadow: '0 4px 5px rgba(0,0,0,.15)',
        zIndex: 999,
        display: this.state.isOpen ? 'block' : 'none'
      };
      var fsResults = {
        margin: '0 4px 4px 0',
        maxHeight: '360px',
        width: 'calc(100% - 4px)',
        padding: '0 0 0 4px',
        position: 'relative',
        overflowX: 'auto',
        overflowY: 'auto'
      };
      var carret: string = this.state.isOpen ? 'ms-Icon ms-Icon--ChevronUp' : 'ms-Icon ms-Icon--ChevronDown';
      var foundSelected = false;
      //Renders content
      return (
        <div style={{ marginBottom: '8px'}}>
          <Label>{this.props.label}</Label>
          <div style={fontSelect}>
            <a style={fontSelectA} onClick={this.onOpenDialog}
              onMouseEnter={this.mouseEnterDropDown} onMouseLeave={this.mouseLeaveDropDown} role="menuitem">
              <span style={fontSelectASpan}>
                {this.state.activeNodes.map((elm: IDropDownTreeViewNode, index?: number) => {
                    if (index !== undefined && index == 0) {
                      return (
                            <span key={this._key + '-spanselect-' + index}>{elm.label}</span>
                      );
                    }
                    else {
                      return (
                            <span key={this._key + '-spanselect-' + index}>, {elm.label}</span>
                      );
                    }
                  })
                }
              </span>
              <div style={fontSelectADiv}>
                <i style={fontSelectADivB} className={carret}></i>
              </div>
            </a>
            <div style={fsDrop}>
              <div style={fsResults}>
                { this.state.tree.map((rootNode: IDropDownTreeViewNode, index: number) => {
                    return (
                      <Tree
                        paddingLeft={this.props.nodesPaddingLeft}
                        tree={rootNode}
                        isNodeCollapsed={false}
                        onChange={this.handleTreeChange.bind(null, rootNode, index)}
                        renderNode={this.renderNode}
                        draggable={false}
                        key={'rootNode-' + index}
                      />
                    );
                  })
                }
              </div>
            </div>
          </div>
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