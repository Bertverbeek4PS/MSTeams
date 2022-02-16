pageextension 60100 MSTeams extends "Customer List"
{

    actions
    {
        addlast(navigation)
        {
            action(msteams)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    JsonResponse: JsonObject;
                    JsonToken: JsonToken;
                    JLinesArray: JsonArray;
                    JmsTeamsToken: JsonToken;
                    JmsTeams: JsonObject;
                    JsonToken1: JsonToken;
                    JsonValue: JsonValue;
                    Team: Text;
                begin
                    JsonResponse := GraphAPIHelper.ListTeams();
                    if JsonResponse.Get('value', JsonToken) then begin
                        JLinesArray := JsonToken.AsArray();

                        foreach JmsTeamsToken in JLinesArray do begin
                            JmsTeams := JmsTeamsToken.AsObject();
                            JmsTeams.Get('id', JsonToken1);
                            JsonValue := JsonToken1.AsValue();
                            Team := JsonValue.AsText();
                        end;
                    end;
                    GraphAPIHelper.SendMessageInChannel('bbbedb03-a292-4064-8aef-900134602b7d', '19:8d51ebe49914414eb183fcf795c4cb5c@thread.tacv2', 'Hello World');
                end;
            }
        }
    }

    var
        GraphAPIHelper: codeunit "Graph API Helper";
}