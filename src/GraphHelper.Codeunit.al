//add calender https://docs.microsoft.com/en-us/graph/api/group-post-events?view=graph-rest-1.0&tabs=http
//add tab as cloud storage
//add chat in channel https://docs.microsoft.com/en-us/graph/api/chatmessage-post?view=graph-rest-1.0&tabs=http
//Use of Json Management??

codeunit 60100 "Graph API Helper"
{
    procedure GetTeam(TeamId: Text): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId;
        Exit(RequestMessage(Uri, Method::Get, Body));
    end;

    procedure PostTeam(DisplayName: Text; Description: Text; OwnerEmail: text): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
        JObjBody: JsonObject;
        JArrMembers: JsonArray;
        JArrRoles: JsonArray;
        JObjmembers: JsonObject;
    begin
        Uri := MSTeamsRootQueryUri + '/teams';

        JObjBody.Add('template@odata.bind', 'https://graph.microsoft.com/v1.0/teamsTemplates(''standard'')');
        JObjBody.Add('displayName', DisplayName);
        JObjBody.Add('description', Description);

        JObjmembers.Add('@odata.type', '#microsoft.graph.aadUserConversationMember');
        JArrRoles.Add('owner');
        JObjmembers.Add('roles', JArrRoles);
        JObjmembers.Add('user@odata.bind', 'https://graph.microsoft.com/v1.0/users(' + '''' + OwnerEmail + '''' + ')');
        JArrMembers.Add(JObjmembers);

        JObjBody.Add('Members', JArrMembers);
        JObjBody.WriteTo(Body);

        Exit(RequestMessage(Uri, Method::Post, Body));
    end;

    procedure ListTeams(): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
    begin
        Uri := MSTeamsRootQueryUri + '/groups?$filter=resourceProvisioningOptions/Any(x:x eq ''Team'')';
        Exit(RequestMessage(Uri, Method::Get, Body));
    end;

    procedure AddMemberTeams(TeamId: Text; MemberEmail: Text; Owner: Boolean): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
        JObjBody: JsonObject;
        JArrRoles: JsonArray;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/members';
        JObjBody.Add('@odata.type', '#microsoft.graph.aadUserConversationMember');
        if Owner then
            JArrRoles.Add('owner');
        JObjBody.Add('roles', JArrRoles);
        JObjBody.Add('user@odata.bind', 'https://graph.microsoft.com/v1.0/users(' + '''' + MemberEmail + '''' + ')');
        JObjBody.WriteTo(Body);

        Exit(RequestMessage(Uri, Method::Post, Body));
    end;

    procedure ListMemberTeams(TeamId: Text): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/members';
        Exit(RequestMessage(Uri, Method::Get, Body));
    end;

    procedure PostChannel(TeamId: Text; DisplayName: Text; Description: Text; OwnerEMail: Text): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
        JObjBody: JsonObject;
        JObjmembers: JsonObject;
        JArrRoles: JsonArray;
        JArrMembers: JsonArray;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/channels';

        JObjBody.Add('@odata.type', '#Microsoft.Graph.channel');
        JObjBody.Add('displayName', DisplayName);
        JObjBody.Add('description', Description);
        if OwnerEmail <> '' then
            JObjBody.Add('membershipType', 'private')
        else
            JObjBody.Add('membershipType', 'standard');

        if OwnerEmail <> '' then begin
            JObjmembers.Add('@odata.type', '#microsoft.graph.aadUserConversationMember');
            JArrRoles.Add('owner');
            JObjmembers.Add('roles', JArrRoles);
            JObjmembers.Add('user@odata.bind', 'https://graph.microsoft.com/v1.0/users(' + '''' + OwnerEmail + '''' + ')');
            JArrMembers.Add(JObjmembers);

            JObjBody.Add('Members', JArrMembers);
        end;

        JObjBody.WriteTo(Body);

        Exit(RequestMessage(Uri, Method::Post, Body));
    end;

    procedure ListChannels(TeamId: Text): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/channels';
        Exit(RequestMessage(Uri, Method::Get, Body));
    end;

    procedure AddMemberChannel(TeamId: Text; ChannelId: Text; MemberEmail: Text; Owner: Boolean): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
        JObjBody: JsonObject;
        JArrRoles: JsonArray;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/channels/' + ChannelId + '/members';
        JObjBody.Add('@odata.type', '#microsoft.graph.aadUserConversationMember');
        if Owner then
            JArrRoles.Add('owner');
        JObjBody.Add('roles', JArrRoles);
        JObjBody.Add('user@odata.bind', 'https://graph.microsoft.com/v1.0/users(' + '''' + MemberEmail + '''' + ')');
        JObjBody.WriteTo(Body);

        Exit(RequestMessage(Uri, Method::Post, Body));
    end;

    procedure ListMemberChannel(TeamId: Text; ChannelId: Text): JsonObject
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/channels/' + ChannelId + '/members';
        Exit(RequestMessage(Uri, Method::Get, Body));
    end;

    procedure SendMessageInChannel(TeamId: Text; ChannelId: Text; Message: Text): JsonObject //test
    var
        Uri: Text;
        Method: Option Get,Post,Patch;
        Body: Text;
        JObjBody: JsonObject;
        JArrBody: JsonArray;
        JObjContent: JsonObject;
    begin
        Uri := MSTeamsRootQueryUri + '/teams/' + TeamId + '/channels/' + ChannelId + '/messages';
        JObjContent.Add('content', message);
        JObjBody.Add('body', JObjContent);
        JObjBody.WriteTo(Body);

        Exit(RequestMessage(Uri, Method::Post, Body));
    end;

    local procedure RequestMessage(Uri: Text; Method: Option Get,Post,Patch; Body: Text): JsonObject
    var
        Client: HttpClient;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        JsonResponse: JsonObject;
        AccessToken: Text;
        JsonContent: Text;
        Content: HttpContent;
        ContentHeaders: HttpHeaders;
    begin
        AccessToken := GetAccessToken();

        RequestMessage.Method(format(Method));
        RequestMessage.SetRequestUri(Uri);

        if Method = Method::Post then begin
            Content.WriteFrom(body);
            Content.GetHeaders(ContentHeaders);
            ContentHeaders.Remove('Content-Type');
            ContentHeaders.Add('Content-Type', 'application/json');
            RequestMessage.Content(Content);
        end;

        Client.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));
        Client.DefaultRequestHeaders().Add('Accept', 'application/json');

        if Client.Send(RequestMessage, ResponseMessage) then begin
            if ResponseMessage.HttpStatusCode = 200 then begin
                ResponseMessage.Content.ReadAs(JsonContent);
                JsonResponse.ReadFrom(JsonContent);
                exit(JsonResponse);
            end;
        end;
    end;

    local procedure GetAccessToken(): Text
    var
        OAuth2: Codeunit OAuth2;
        MsTeamsCredentials: Codeunit "MS Teams Credentials";
        MSTeamsSetup: Record "MS Teams Setup";
        AccessToken: Text;
        AuthCodeError: Text;
        Scopes: List of [Text];
    begin
        MSTeamsSetup.Get;

        if not MsTeamsCredentials.IsInitialized() then
            MsTeamsCredentials.Init();

        Scopes.Add('https://graph.microsoft.com/.default');
        OAuth2.AcquireTokenWithClientCredentials(
            MsTeamsCredentials.GetClientID(),
            MsTeamsCredentials.GetClientSecret(),
            MSTeamsSetup."OAuth Authority Url",
            MSTeamsSetup."Redirect URL",
            Scopes,
            AccessToken);

        if (AccessToken = '') or (AuthCodeError <> '') then
            Error(AuthCodeError);

        exit(AccessToken);
    end;

    local procedure MSTeamsRootQueryUri(): Text
    var
        MSTeamsSetup: Record "MS Teams Setup";
    begin
        If MSTeamsSetup.Get then
            exit(MSTeamsSetup."MS Teams Root Query URL");
    end;
}