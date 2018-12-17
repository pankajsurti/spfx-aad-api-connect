import { ClientMode } from "./ClientMode";
import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface ISampleMsGraphApiProps {
  description: string;
  clientMode: ClientMode;
  context: WebPartContext;
}
