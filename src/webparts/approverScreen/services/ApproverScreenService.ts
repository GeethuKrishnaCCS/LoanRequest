import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";

export class ApproverScreenService extends BaseService {
    private _spfi: SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this._spfi = getSP(context);
    }
    public getListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }

    public EnsureUser(username: string) {
        return this._spfi.web.ensureUser(username);
    }

    // public addListItem(data: any, listname: string, url: string): Promise<any> {
    //     return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    // } 

    // public getFloorListItems(listname: string, id: number, url: string): Promise<any> {
    //     return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Floor,ID,RepaymentType/Title,RepaymentType/ID").expand("RepaymentType").filter("RepaymentType/ID eq '" + id + "'")();
    // }
    
    // public getRepaymentRequestListItems(listname: string, id: number, url: string): Promise<any> {
    //     return this._spfi.web.getList(`${url}/Lists/${listname}`).items
    //         .select("LoanType", "ID", "RepaymentType/RepaymentType", "RepaymentType/ID")
    //         .expand("RepaymentType")
    //         .filter(`ID eq ${id}`) // Optional: to filter by specific ID if needed
    //         ();
    // }
    

    /*public getFloorListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Floor,ID,Building/Title,Building/ID").expand("Building").filter("Building/ID eq '" + id + "'")();
    }*/

    // public getProjectListItems(listname: string, id: number, url: string): Promise<any> {
    //     return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Project,ID,Program/Title,Program/ID").expand("Program").filter("Program/ID eq '" + id + "'")();
    // }
   

    // public async getUser(userId: number): Promise<any> {
    //     return this._spfi.web.getUserById(userId)();
    // }

    // public getItemSelectExpandFilter(siteUrl: string, listname: string, select: string, expand: string, filter: string): Promise<any> {
    //     return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
    //         .select(select)
    //         .expand(expand)
    //         .filter(filter)()
    // }
    // public getItemSelectExpand(siteUrl: string, listname: string, select: string, expand: string): Promise<any> {
    //     return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
    //         .select(select)
    //         .expand(expand)
    //         ()
    // }

    
}