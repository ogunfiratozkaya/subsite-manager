export interface IUser {
    Id: number;
    FullName: string;
    EMail: string;
}
export default class User {
    public Id: number;
    public FullName: string;
    public EMail: string;
    public Image: string;

    constructor(options: IUser) {
        this.Id = options.Id;
        this.FullName = options.FullName;
        this.EMail = options.EMail;

        this.Image = `/_layouts/15/userphoto.aspx?size=L&accountname=${options.EMail}`;

    }
}