import * as React from 'react';
import styles from './VisioOnlineReact.module.scss';
import { IVisioOnlineReactProps } from './IVisioOnlineReactProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class VisioOnlineReact extends React.Component<IVisioOnlineReactProps, {}> {


  constructor(props: IVisioOnlineReactProps) {
    super(props);

    // set delegate functions that will be used to pass the values from the Visio service to the component
    // this.props.visioService.onSelectionChanged = this._onSelectionChanged;
    this.props.visioService._isControlKeyPressed = false;
    // this.state = {
    //   isControlKeyPressed: false
    // };
  }


  public render(): React.ReactElement<IVisioOnlineReactProps> {
    return (
      <div className={styles.visioOnlineReact}>
        <div id='iframeHost' className={styles.iframeHost}></div>
      </div>
    );
  }
  
    /**
   * method executed after a on selection change event is triggered
   * @param selectedShape the shape selected by the user on the Visio diagram
   */
  // private _onSelectionChanged = (selectedShape: Visio.Shape): void => {

  //   console.log("Selected shape: ", selectedShape);
  //   console.log("Selected shape name: ", selectedShape.name);

  //   this.setState({
  //     isControlKeyPressed: true
  //   });
  // }

  public handleKeyPressDown(e) {
    this.state= {currentKey: e.keyCode};
    console.log('You just pressed a key!');
        if(e.keyCode === 17) {
      console.log('You just pressed Control!');
      this.props.visioService._isControlKeyPressed = true;
    }
  }

  public handleKeyPressUp(e) {
    this.state= {currentKey: e.keyCode};
    console.log('You just released a key!');
        if(e.keyCode === 17) {
      console.log('You just released Control!');
      this.props.visioService._isControlKeyPressed = true;
    }
  }

  public componentWillUnmount() {
    document.removeEventListener('keydown', this.handleKeyPressDown);
    document.removeEventListener('keyup', this.handleKeyPressUp);
  }
  public componentDidMount() {
    document.addEventListener('keydown', this.handleKeyPressDown);
    document.addEventListener('keyup', this.handleKeyPressUp);
        if (this.props.documentUrl) {
      this.props.visioService.load(this.props.documentUrl, this.props.zoomLevel);
    }
  }

  public async componentDidUpdate(prevProps: IVisioOnlineReactProps) {
    console.log("componentDidUpdate function");
    if (this.props.documentUrl && this.props.documentUrl !== prevProps.documentUrl) {
      this.props.visioService.load(this.props.documentUrl);
    }
    if (this.props.showShapeNameFlyout !== prevProps.showShapeNameFlyout) {
      this.props.visioService.Options(this.props.showShapeNameFlyout);
    }
    if ((this.props.bHighLight !== prevProps.bHighLight) || (this.props.shapeName !== prevProps.shapeName)) {
      this.props.visioService.highlightShape(this.props.shapeName, this.props.bHighLight);
    }
    if ((this.props.bOverlay !== prevProps.bOverlay) || (this.props.shapeName !== prevProps.shapeName)) {
      this.props.visioService.addOverlay(this.props.shapeName, this.props.bOverlay, this.props.overlayType,this.props.overlayText,
        this.props.overlayWidth,this.props.overlayHeight);
    }
  }

}
