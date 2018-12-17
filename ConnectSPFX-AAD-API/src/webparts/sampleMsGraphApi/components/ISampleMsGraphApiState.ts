import { IUserItem } from './IUserItem';

export interface ISampleMsGraphApiState {
   users: Array<IUserItem>;
   searchFor: string;
 }