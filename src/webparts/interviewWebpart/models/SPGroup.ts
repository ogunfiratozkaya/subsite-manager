export interface ISPGroup {
    Id: number;
    Title: string;
    Url: string;
}
export default class SPGroup {
    public Id: number;
    public Title: string;
    public Url: string;
    public Owners: string;
    public Members: string;
    public Visitors: string;

    constructor(options: ISPGroup) {
        this.Id = options.Id;
        this.Title = options.Title;
        this.Url = options.Url;
        this.Owners = `${options.Title} Owners`;
        this.Members = `${options.Title} Members`;
        this.Visitors = `${options.Title} Visitors`;
    }
}