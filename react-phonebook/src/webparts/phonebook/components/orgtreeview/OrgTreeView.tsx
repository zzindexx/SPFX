import * as React from 'react';
import Tree, {TreeNode} from 'rc-tree'
import 'rc-tree/assets/index.css';
import { IOrgTreeViewProps } from './IOrgTreeViewProps';
import { IOrgTreeViewState } from './IOrgTreeViewState';
import { IOrgUnit } from '../../../../classes/IOrgUnit';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';





export default class OrgTreeView extends React.Component<IOrgTreeViewProps, IOrgTreeViewState> {
  constructor(props: IOrgTreeViewProps) {
    super(props);

    this.state = {
      isLoading: false,
      units: []
    };
    
  }

  public async componentDidMount() {
    this.setState({isLoading:true, units: []});

    let orgunits: IOrgUnit[] = await this.props.getData();
    this.setState({isLoading:false, units: orgunits});
  }

  private generateTreeNodes(node:IOrgUnit){
    var result = {title: node.title, children: []};
    if (node.child_orgunits.length>0){
      node.child_orgunits.forEach((unit:IOrgUnit) => {
        result.children.push(this.generateTreeNodes(unit));
      });
    }
    return result;
  }

  public render(): React.ReactElement<IOrgTreeViewProps> {
    const isLoading: JSX.Element = this.state.isLoading ? <div><Spinner label="Loading" /></div>: <div></div>;

    let treenodes = [];
    this.state.units.forEach((unit:IOrgUnit)=> {
      treenodes.push(this.generateTreeNodes(unit));
    });

    return (
      <div>
        {isLoading}
        <Tree treeData={treenodes}>
        </Tree>
      </div>
    );
  }
}
