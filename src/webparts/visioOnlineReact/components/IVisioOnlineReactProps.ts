import { VisioService } from "../../../shared/services/VisioService";

export interface IVisioOnlineReactProps {
  description: string;
  documentUrl: string;
  zoomLevel: string;
  shapeName:string;
  showShapeNameFlyout: boolean;
  bHighLight:boolean;
  bOverlay:boolean;
  overlayType:string;
  overlayText:string;
  overlayWidth:string;
  overlayHeight:string;
  visioService:VisioService;
}
