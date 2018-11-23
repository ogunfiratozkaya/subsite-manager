
import { Web, WebAddResult, sp, SPBatch } from "@pnp/sp";
import { GroupAddResult } from "@pnp/sp/src/sitegroups";
import Constants from "../models/Constants";
import SPGroup from "../models/SPGroup";
import User from "../models/User";

export class RoleAssignment {
    public principalId: number;
    public roleDefId: number;

    constructor(principalId, roleDefId) {
        this.roleDefId = roleDefId;
        this.principalId = principalId;
    }
}

export class SharePointService {

    public CreateSharePointGroup(web: Web, title: string, description: string) {
        return new Promise((resolve, reject) => {
            web.siteGroups.add({
                Title: title,
                Description: description,
                AllowRequestToJoinLeave: true,
                AutoAcceptRequestToJoinLeave: true,
                AllowMembersEditMembership: true,
            }).then((w: GroupAddResult) => {
                // show the response from the server when adding the web
                console.log(w.data);
                resolve(w);
            });
        });
    }

    public AddPermissionToWeb(web: Web, roleAssignments: RoleAssignment[]) {
        return new Promise((resolve, reject) => {
            const spQueue = {
                counter: roleAssignments.length,
                success: () => {
                    spQueue.counter--;
                    if (spQueue.counter == 0) {
                        resolve();
                    }
                }
            };
            for (let index = 0; index < roleAssignments.length; index++) {
                const roleAssignment = roleAssignments[index];
                web.roleAssignments.add(roleAssignment.principalId, roleAssignment.roleDefId)
                    .then((w) => {
                        spQueue.success();
                        // show the response from the server when adding the web
                    });
            }

        });
    }


    public CreateSharePointGroups(web: Web, groups: string[]) {
        return new Promise((resolve, reject) => {
            const result = new Array<GroupAddResult>();
            const spQueue = {
                counter: groups.length,
                success: () => {
                    spQueue.counter--;
                    if (spQueue.counter == 0) {
                        resolve(result);
                    }
                }
            };
            for (let index = 0; index < groups.length; index++) {
                const title = groups[index];
                web.siteGroups.add({
                    Title: title,
                    AllowRequestToJoinLeave: true,
                    AutoAcceptRequestToJoinLeave: true,
                    AllowMembersEditMembership: true,
                }).then((group: GroupAddResult) => {
                    result.push(group);
                    spQueue.success();
                    // show the response from the server when adding the web
                });
            }
        });
    }



    public _CreateSubSite(webUrl, title, urlName, description, template, language, inheritancePermission) {
        const web = new Web(webUrl);
        return new Promise((resolve, reject) => {
            web.webs.add(title, urlName, description, template, language, inheritancePermission).then((w: WebAddResult) => {
                // show the response from the server when adding the web
                console.log(w.data);
                resolve(w);
            });
        });
    }

    public CreateSubSite(webUrl, title, urlName, description, template, language, inheritancePermission) {
        return new Promise((resolve, reject) => {
            this._CreateSubSite(webUrl, title, urlName, description, template, language, inheritancePermission).then((w: WebAddResult) => {
                console.log(w.data);
                const ownersGroupName = `${w.data.Title} Owners`;
                const membersGroupName = `${w.data.Title} Members`;
                const visitorsGroupName = `${w.data.Title} Visitors`;

                this.CreateSharePointGroups(w.web, [ownersGroupName, membersGroupName, visitorsGroupName])
                    .then((groups: GroupAddResult[]) => {
                        const roleAssignments = new Array<RoleAssignment>();
                        for (let index = 0; index < groups.length; index++) {
                            const groupTitle = groups[index].data.Title;
                            const groupId = groups[index].data.Id;

                            switch (groupTitle) {
                                case ownersGroupName:
                                    roleAssignments.push(new RoleAssignment(
                                        groupId,
                                        Constants.FULL_CONTROL_DEFINITION_ID
                                    ));
                                    break;
                                case membersGroupName:
                                    roleAssignments.push(new RoleAssignment(
                                        groupId,
                                        Constants.CONTRIBUTE_DEFINITION_ID
                                    ));
                                    break;
                                case visitorsGroupName:
                                    roleAssignments.push(new RoleAssignment(
                                        groupId,
                                        Constants.READ_DEFINITION_ID
                                    ));
                                    break;
                            }
                        }

                        this.AddPermissionToWeb(w.web, roleAssignments).then(() => {
                            resolve([w.web, groups]);
                        });
                    });

            });
        });
    }

    public GetMembersByGroupName(webUrl: string, groupTitle: string): Promise<User[]> {
        return new Promise((resolve, reject) => {
            const web = new Web(webUrl);
            web.siteGroups
                .getByName(groupTitle)
                .users
                .get().then((users) => {
                    const result = users.map((user) => {
                        return new User({
                            EMail: user.Email,
                            FullName: user.Title,
                            Id: user.Id
                        });
                    });

                    resolve(result);
                })
                .catch((err) => {
                    debugger;
                });
        });
    }

    public GetAllSites(webUrl: string): Promise<SPGroup> {
        return new Promise((resolve, reject) => {
            const web = new Web(webUrl);

            web.webs
                .select("Title, Url,Id")
                .get()
                .then((items: any) => {
                    // show the response from the server when adding the web
                    resolve(items.map((item) => {
                        return new SPGroup({
                            Id: item.Id,
                            Title: item.Title,
                            Url: item.Url
                        });
                    }));
                }).catch((err) => {
                    reject(err);
                });
        });
    }
}


export default new SharePointService();