class MSALConfig {
    static readonly appId :string=  '9897d020-9689-4479-a310-ce46937d6816';
    static readonly scopes :Array<string> = [
      "user.read",
      "calendars.read",
      "Sites.ReadWrite.All"
    ];
}
 
export default MSALConfig;