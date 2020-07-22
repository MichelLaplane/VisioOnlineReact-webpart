import { VisioService } from "../../../shared/services/VisioService";

export interface IVisioOnlineReactProps {
  description: string;
  documentUrl: string;
  zoomLevel: string;
  shapeName: string;
  bHighLight: boolean;
  bOverlay: boolean;
  visioService: VisioService;
}
