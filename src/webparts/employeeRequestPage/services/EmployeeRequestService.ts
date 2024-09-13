import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";

export class EmployeeRequestService extends BaseService {
    private _spfi: SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this._spfi = getSP(context);
    }

    // for get list item

    public getListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }

    public async getListfilter(listname: string,designation: string, url: string): Promise<any> {
        return await this._spfi.web.getList(url + "/Lists/" + listname).items.filter(`Designation eq '${designation}'`).getAll();
    }
    public async getMaxAmountfilter(listname: string,id: number, url: string): Promise<any> {
        return await this._spfi.web.getList(url + "/Lists/" + listname).items.filter("JobBand/ID eq '" + id + "'").getAll();
    }

    public async getjobbandidfrommasterlist(listname: string,id: number, url: string): Promise<any> {
        return await this._spfi.web.getList(url + "/Lists/" + listname).items.select("*,ID,LoanType/ID, LoanType/LoanType,JobBand/Bands,JobBand/ID").expand("JobBand, LoanType").filter("JobBand/ID eq '" + id + "'")();
    }

    public async getSelectExpandFilter(listname: string, url: string, select: string, expand: string,filter: string): Promise<any> {
        return await this._spfi.web.getList(url + "/Lists/" + listname).items.select(select).expand(expand).filter(filter)();
    }

    public addListItem(data: any, listname: string, url: string): Promise<any> {

        console.log(data, "DATA") ;
        console.log(listname, "LISTNAME");
        console.log(url,URL)
        return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    }

    // public addListItem(data: any, listname: string, url: string): Promise<any> {
    //     console.log(data, "DATA") ;
    //     console.log(listname, "LISTNAME");
    //     console.log(url,URL)
    //     return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    // } 


    // public getFloorListItems(listname: string, id: number, url: string): Promise<any> {
    //     return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Floor,ID,RepaymentType/Title,RepaymentType/ID").expand("RepaymentType").filter("RepaymentType/ID eq '" + id + "'")();
    // }

    public getRepaymentRequestListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(`${url}/Lists/${listname}`).items
            .select("LoanType", "ID", "RepaymentType/RepaymentType", "RepaymentType/ID")
            .expand("RepaymentType")
            .filter(`ID eq ${id}`) // Optional: to filter by specific ID if needed
            ();
    }

    public getChoiceListItems(url: string, listname: string, field: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).fields.getByInternalNameOrTitle(field)();
    }

    public EnsureUser(username: string) {
        return this._spfi.web.ensureUser(username);
    }

    public addItemRequestForm(data: any, listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    }

    public updateRequestForm(listname: string, data: any, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
 


    /*public getFloorListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Floor,ID,Building/Title,Building/ID").expand("Building").filter("Building/ID eq '" + id + "'")();
    }*/

    // public getProjectListItems(listname: string, id: number, url: string): Promise<any> {
    //     return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Project,ID,Program/Title,Program/ID").expand("Program").filter("Program/ID eq '" + id + "'")();
    // }


    public async getUser(userId: number): Promise<any> {
        return this._spfi.web.getUserById(userId)();
    }

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