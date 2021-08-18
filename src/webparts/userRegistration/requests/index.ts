import { sp, SPHttpClient } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import axios from 'axios';

const client = new SPHttpClient();

export const getSiteGroups = () => {
    return (
        sp.web.siteGroups()
            .then(res => {
                return res
            })
    )
}

export const addUserToGroup = (siteUrl, email, id) => {
    const client = new SPHttpClient();
    return (
        client.post(`${siteUrl}/_api/SP.Web.ShareObject`, {
            body: JSON.stringify({
                emailBody: 'Welcome to site',
                includeAnonymousLinkInEmail: false,
                peoplePickerInput: JSON.stringify([{
                    Key: email,
                    DisplayText: email,
                    IsResolved: true,
                    Description: email,
                    EntityType: '',
                    EntityData: {
                        SPUserID: email,
                        Email: email,
                        IsBlocked: 'False',
                        PrincipalType: 'UNVALIDATED_EMAIL_ADDRESS',
                        AccountName: email,
                        SIPAddress: email,
                        IsBlockedOnODB: 'False'
                    },
                    MultipleMatches: [],
                    ProviderName: '',
                    ProviderDisplayName: ''
                }]),
                roleValue: `group:${id}}`, // where `6` is a GroupId
                sendEmail: true,
                url: siteUrl,
                useSimplifiedRoles: true
            })
        })
            .then(r => r.json())
            .then((res) => {
                return res;
            })
    )
}

export const shareSiteWithUser = (email, id) => {
    return (
        sp.web.siteGroups.getById(id).users
            .add(`i:0#.f|membership|${email}`).then(function (d) {
                console.log(d)
                return d;
            })
    )
}

export const updateUserByID = ( id, data ) => {
    return (
        axios({
            method: 'post',
            headers: {
                'app-secret': '',
            },
            url: '/user/update',
            data: data
          })
          .then(res => res)
    )
}