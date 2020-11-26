// Original Code from Joel Rodrigues
// https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/react-visio
// Modified by Michel Laplane
// https://github.com/MichelLaplane/VisioOnlineReact-webpart
import { find } from '@microsoft/sp-lodash-subset';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Dictionary } from 'lodash';

export class VisioService {

  private _webPartContext: IWebPartContext;

  private _url = "";
  private _zoomLevel: string = "100";
  /**
   * gets the url of the Visio document to embed
   * @returns a string with the document url
   */
  get url(): string {
    return this._url;
  }
  /**
   * sets the url of the Visio document to embed
   * @param url the url of the document
   */
  set url(url: string) {
    // apis are enabled for EmbedView action only
    url = url.replace("action=view", "action=embedview");
    url = url.replace("action=interactivepreview", "action=embedview");
    url = url.replace("action=default", "action=embedview");
    url = url.replace("action=edit", "action=embedview");

    this._url = url;
  }

  private _session: OfficeExtension.EmbeddedSession = null;
  private _shapes: Visio.Shape[] = [];
  private _selectedShape: Visio.Shape;

  private _overlayedShape: Dictionary<string> = {};

  private _documentLoadComplete = false;
  private _pageLoadComplete = false;

  private _foundedShape: Visio.Shape;
  private _enteredShape: Visio.Shape;
  private _leavedShape: Visio.Shape;

  private _bShowShapeNameFlyout: boolean;

  public _isControlKeyPressed:boolean;

  /**
   * gets a pre-loaded collection of relevant shapes from the diagram
   */
  public get shapes(): Visio.Shape[] {
    return this._shapes;
  }

  // delegate functions passed from the react component
  public onSelectionChanged: (selectedShape: Visio.Shape) => void;
  public getAllShapes: (shapes: Visio.Shape[]) => void;

  /**
   * class constructor
   * @param webPartContext the context of the web part
   */
  constructor(webPartContext: IWebPartContext) {
    // set web part context
    this._webPartContext = webPartContext;
  }

  /**
 * initializes the embed session and attaches event handlers
 * this is the function that should be called to start the session
 * @param docUrl embed url of the document
 * @returns returns a promise
 */
  public load = async (docUrl: string, zoomLevel?: string): Promise<void> => {
    console.log("Start loading Visio data");

    try {

      // sets the url, modifying it if required - uses set method to re-use logic
      this.url = docUrl;
      if (zoomLevel != undefined)
        this._zoomLevel = zoomLevel;
      // init
      await this._init();

      // add custom onDocumentLoadComplete event handler
      await this._addCustomEventHandlers();

      // trigger document and page loaded event handlers after 3 seconds in case Visio fails to trigger them
      // this is randomly happening on chrome, but seems to always fail on IE...
      setTimeout(() => {
        this._onDocumentLoadComplete(null);
        this._onPageLoadComplete(null);
      }, 3000);

    } catch (error) {
      this.logError(error);
    }
  }

  /**
   * initialize session by embedding the Visio diagram on the page
   * @returns returns a promise
   */
  private async _init(): Promise<any> {
    // Remove element for preventing to have multiple iFrame
    let docRootElement = document.getElementById("iframeHost");
    if (docRootElement != null) {
      for (var i = 0; i < docRootElement.childNodes.length; ++i) {
        docRootElement.removeChild(docRootElement.childNodes[i]);
      }
      console.log("Document Url = " + this._url);
      console.log("Document Url with wdzoom = " + this._url + '&wdzoom=' + this._zoomLevel);
      // initialize communication between the developer frame and the Visio Online frame
      this._session = new OfficeExtension.EmbeddedSession(
        this._url + '&wdzoom=' + this._zoomLevel, {
        id: "embed-iframe",
        container: document.getElementById("iframeHost"),
        width: "100%",
        height: "600px"
      }
      );
      await this._session.init();
      console.log("Session successfully initialized");
    }
  }

  /**
   * function to add custom event handlers
   * @returns returns a promise
   */
  private _addCustomEventHandlers = async (): Promise<any> => {

    try {
      await Visio.run(this._session, async (context: Visio.RequestContext) => {
        var doc: Visio.Document = context.document;

        // on document load complete
        const onDocumentLoadCompleteEventResult: OfficeExtension.EventHandlerResult<Visio.DocumentLoadCompleteEventArgs> =
          doc.onDocumentLoadComplete.add(
            this._onDocumentLoadComplete
          );
        // on page load complete
        const onPageLoadCompleteEventResult: OfficeExtension.EventHandlerResult<Visio.PageLoadCompleteEventArgs> =
          doc.onPageLoadComplete.add(
            this._onPageLoadComplete
          );
        // on selection changed
        const onSelectionChangedEventResult: OfficeExtension.EventHandlerResult<Visio.SelectionChangedEventArgs> =
          doc.onSelectionChanged.add(
            this._onSelectionChanged
          );
        // on mouse enter
        const onShapeMouseEnterEventResult: OfficeExtension.EventHandlerResult<Visio.ShapeMouseEnterEventArgs> =
          doc.onShapeMouseEnter.add(
            this._onShapeMouseEnter
          );
        // on mouse leave
        const onShapeMouseLeaveEventResult: OfficeExtension.EventHandlerResult<Visio.ShapeMouseEnterEventArgs> =
          doc.onShapeMouseLeave.add(
            this._onShapeMouseLeave
          );
        await context.sync();
        console.log("Document Load Complete handler attached");
      });
    } catch (error) {
      this.logError(error);
    }
  }

  /**
   * method executed after a on document load complete event is triggered
   * @param args event arguments
   * @returns returns a promise
   */
  private _onDocumentLoadComplete = async (args: Visio.DocumentLoadCompleteEventArgs): Promise<void> => {

    // only execute if not executed yet
    if (!this._documentLoadComplete) {

      try {
        console.log("Document Loaded Event: " + JSON.stringify(args));

        // set internal flag to prevent event from running again if triggered twice
        this._documentLoadComplete = true;

        await Visio.run(this._session, async (context: Visio.RequestContext) => {
          var doc: Visio.Document = context.document;
          // disable Hyperlinks on embed diagram
          doc.view.disableHyperlinks = true;
          // hide diagram boundary on embed diagram
          doc.view.hideDiagramBoundary = true;

          await context.sync();
        });


      } catch (error) {
        this.logError(error);
      }
    }
  }

  /**
   * method executed after a on page load event is triggered
   * @param args event arguments
   * @returns returns a promise
   */
  private _onPageLoadComplete = async (args: Visio.PageLoadCompleteEventArgs): Promise<void> => {

    // only execute if not executed yet
    if (!this._pageLoadComplete) {

      try {
        console.log("Page Loaded Event: " + JSON.stringify(args));

        // set internal flag to prevent event from running again if triggered twice
        this._pageLoadComplete = true;

        // get all relevant shapes and populate the class variable
        this._shapes = await this._getAllShapes();

        // call delegate function from the react component
        this.getAllShapes(this._shapes);

      } catch (error) {
        this.logError(error);
      }
    }
  }

  /**
   * get all shapes from page
   * @returns returns a promise
   */
  private _getAllShapes = async (): Promise<Visio.Shape[]> => {

    console.log("Getting all shapes");

    try {
      let shapes: Visio.Shape[];

      await Visio.run(this._session, async (context: Visio.RequestContext) => {
        const page: Visio.Page = context.document.getActivePage();
        const shapesCollection: Visio.ShapeCollection = page.shapes;
        shapesCollection.load();
        await context.sync();

        // load all required properties for each shape
        for (let i: number = 0; i < shapesCollection.items.length; i++) {
          shapesCollection.items[i].shapeDataItems.load();
          shapesCollection.items[i].hyperlinks.load();
        }
        await context.sync();

        shapes = shapesCollection.items;

        return shapes;
      });

      return shapes;
    } catch (error) {
      this.logError(error);
    }
  }

  /**
   * method executed after a on selection change event is triggered
   * @param args event arguments
   * @returns returns a promise
   */
  private _onSelectionChanged = async (args: Visio.SelectionChangedEventArgs): Promise<void> => {

    try {
      console.log("Selection Changed Event " + JSON.stringify(args));

      if (args.shapeNames.length > 0 && this._shapes && this._shapes.length > 0) {

        // get name of selected item
        const selectedShapeText: string = args.shapeNames[0];

        // find selected shape on the list of pre-loaded shapes
        this._selectedShape = find(this._shapes,
          s => s.name === selectedShapeText
        );

        // call delegate function from the react component
        this.onSelectionChanged(this._selectedShape);

      } else {
        // shape was deselected
        this._selectedShape = null;
      }
    } catch (error) {
      this.logError(error);
    }
  }

  /**
   * select a shape on the visio diagram
   * @param name the name of the shape to select
   */
  public selectShape = async (name: string): Promise<void> => {

    try {

      // find the correct shape from the pre-loaded list of shapes
      // check the ShapeData item with the 'Name' key
      const shape: Visio.Shape = find(this._shapes,
        s => (find(s.shapeDataItems.items, i => i.label === "Name").value === name)
      );

      // only select shape if not the currently selected one
      if (this._selectedShape === null
        || this._selectedShape === undefined
        || (this._selectedShape && this._selectedShape.name !== shape.name)) {

        await Visio.run(this._session, async (context: Visio.RequestContext) => {
          const page: Visio.Page = context.document.getActivePage();
          const shapesCollection: Visio.ShapeCollection = page.shapes;
          shapesCollection.load();
          await context.sync();

          const diagramShape: Visio.Shape = shapesCollection.getItem(shape.name);
          // select shape on diagram
          diagramShape.select = true;

          await context.sync();
          console.log(`Selected shape '${shape.name}' in diagram`);
          this._selectedShape = shape;
        });
      } else {
        console.log(`Shape '${shape.name}' is already selected in diagram`);
      }

    } catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // ShapeMouseEnter event
  // @param args
  // @returns returns a promise
  // ========================================================
  private _onShapeMouseEnter = async (args: Visio.ShapeMouseEnterEventArgs): Promise<void> => {
    try {
      if (this._bShowShapeNameFlyout == true) {
        const enteredShapeName: string = args.shapeName;
        console.log("_onShapeMouseEnter enteredShapeName = " + enteredShapeName);
        await this.getShapeFromName(enteredShapeName);
        this._enteredShape = this._foundedShape;
        console.log("Entered Shape = " + this._enteredShape.name);
        await this.addOverlay(this._enteredShape.name, true, "Text", enteredShapeName);
        console.log("_onShapeMouseEnter end");
      }
    }
    catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // ShapeMouseLeave event
  // @param args
  // @returns returns a promise
  // ========================================================
  private _onShapeMouseLeave = async (args: Visio.ShapeMouseLeaveEventArgs): Promise<void> => {
    try {
      if (this._bShowShapeNameFlyout == true) {
        const leavedShapeName: string = args.shapeName;
        console.log("_onShapeMouseLeave leavedShapeName = " + leavedShapeName);
        await this.getShapeFromName(leavedShapeName);
        this._leavedShape = this._foundedShape;
        console.log("Leaved Shape = " + this._leavedShape.name);
        await this.addOverlay(this._leavedShape.name, false, "Text", leavedShapeName);
        console.log("_onShapeMouseLeave end");
      }
    }
    catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // Get a shape by its name
  // @param shapeName name of the shape to retrieve
  // @returns returns a promise
  // ========================================================
  private getShapeFromName = async (shapeName: string): Promise<any> => {
    try {
      console.log("getShapeFromName entrée");
      var shapeID: any = shapeName.substring(shapeName.lastIndexOf(".") + 1);
      await Visio.run(this._session, async (context: Visio.RequestContext) => {
        const activePage: Visio.Page = context.document.getActivePage();
        const shapesCollection: Visio.ShapeCollection = activePage.shapes;
        shapesCollection.load();
        await context.sync();
        //        const shape: Visio.Shape = shapesCollection.items[shapeID - 1];
        const shape: Visio.Shape = shapesCollection.getItem(shapeName);
        shape.load();
        this._foundedShape = shape;
        console.log("_getShapeFromName sortie");
      });
    }
    catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // Set options of the service
  // @param wether to display the Shapename flyout or not
  // @returns returns a promise
  // ========================================================
  public Options = async (bShowShapeNameFlyout: boolean): Promise<void> => {
    this._bShowShapeNameFlyout = bShowShapeNameFlyout;
  }

  // ========================================================
  // Highlight a Shape
  // @param shapeName name of the shape to higlight
  // @param bHighlight higlight if true un-highlight if false
  // @returns returns a promise
  // ========================================================
  public highlightShape = async (shapeName: string, bHighlight: boolean): Promise<void> => {
    console.log("Start highlightShape : " + shapeName);

    try {
      await Visio.run(this._session, async (context: Visio.RequestContext) => {
        const activePage: Visio.Page = context.document.getActivePage();
        const shapesCollection: Visio.ShapeCollection = activePage.shapes;
        shapesCollection.load();
        await context.sync();
        console.log("shapesCollection.load");
        const shape: Visio.Shape = shapesCollection.getItem(shapeName);        
        shape.load();
        await context.sync();
        console.log("Shape founded : " + shape.name);
        if (bHighlight == true)
          shape.view.highlight = { color: "#FF0000", width: 2 };
        else
          shape.view.highlight = null;
        await context.sync();
      });
    } catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // Add an image overlay to a Shape
  // @param shapeName name of the shape to higlight
  // @param bHighlight higlight if true un-highlight if false
  // @returns returns a promise
  // ========================================================
  public addOverlay = async (shapeName: string, bAddOverlay: boolean, overlayType: string, strOverlay?: string,
    strWidth?: string, strHeight?: string): Promise<void> => {
    var overlayId;

    try {
      var strHtml = ((strOverlay != "") && (strOverlay != undefined)) ? strOverlay : this.getHtmlFlyOut();
      var strText = ((strOverlay != "") && (strOverlay != undefined)) ? strOverlay : this.getTextFlyOut();
      var strImage = ((strOverlay != "") && (strOverlay != undefined)) ? strOverlay : this.getImageFlyOut();
      var width = ((strWidth != "") && (strWidth != undefined)) ? parseInt(strWidth) : 50;
      var height = ((strHeight != "") && (strHeight != undefined)) ? parseInt(strHeight) : 50;
      await Visio.run(this._session, async (context: Visio.RequestContext) => {
        const activePage: Visio.Page = context.document.getActivePage();
        const shapesCollection: Visio.ShapeCollection = activePage.shapes;
        shapesCollection.load();
        await context.sync();
        const shape: Visio.Shape = shapesCollection.getItem(shapeName);
        shape.load();
        if (bAddOverlay) {
          switch (overlayType) {
            case "Text":
              console.log("strText : " + strText);
              overlayId = shape.view.addOverlay("Text", strText, "Center", "Middle", width, height);
              break;
            case "Image":
              console.log("strImage : " + strImage);
              overlayId = shape.view.addOverlay("Image", strImage, "Center", "Middle", width, height);
              break;
            case "Html":
              console.log("strHtml : " + strHtml);
              overlayId = shape.view.addOverlay("Html", strHtml, "Center", "Middle", width, height);
              break;
            default:
              overlayId = shape.view.addOverlay("Text", strText, "Center", "Middle", width, height);
              break;
          }
          await context.sync();
          console.log("overlayId : " + overlayId.value.toString());
          this._overlayedShape[shape.name] = overlayId.value.toString();
        }
        else {
          await context.sync();
          var strId = this._overlayedShape[shape.name];
          console.log("strId : " + strId);
          shape.view.removeOverlay(parseInt(strId));
          await context.sync();
          delete this._overlayedShape[shape.name];
        }
      });
    } catch (error) {
      this.logError(error);
    }
  }


  private getHtmlFlyOut = (): string => {
    var retVal = "";
    retVal = "https://www.bing.com/";
    return retVal;
  }

  private getTextFlyOut = (): string => {
    var retVal = "My text";
    return retVal;
  }

  private getImageFlyOut = (): string => {
    var retVal = "https://www.microsoft.com/favicon.ico?v2";
    return retVal;
  }    

  /**
   * generate embed url for a document
   * @param docId the list item ID of the target document
   */
  private generateEmbedUrl = async (itemProperties: any): Promise<string> => {
    let url: string = "";

    try {
      // check if data was returned
      if (itemProperties) {
        // generate required URL
        const siteUrl: string = this._webPartContext.pageContext.site.absoluteUrl;
        const sourceDoc: string = encodeURIComponent(itemProperties.File.ContentTag.split(",")[0]);
        const fileName: string = encodeURIComponent(itemProperties.File.Name);

        if (siteUrl && sourceDoc && fileName) {
          url = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${sourceDoc}&file=${fileName}&action=default`;
        }
      }

    } catch (error) {
      console.error(error);
    }

    return url;
  }

  /**
   * log error
   * @param error error object
   */
  private logError = (error: any): void => {
    console.error("Error");
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: ", JSON.stringify(error.debugInfo));
    } else {
      console.error(error);
    }
  }
}
