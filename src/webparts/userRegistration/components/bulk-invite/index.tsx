import * as React from 'react';
import CSVReader from 'react-csv-reader'
import { shareSiteWithUser } from '../../requests';

type BulkInviteProps = {
    inviteUser: any;
    updateUser: any;
};
const BulkInvite = (props: BulkInviteProps): JSX.Element => {
    const csvJSON = (csv: any) => {
        var lines = csv;
        var result = [];
        var headers = lines[0];
        for (var i = 1; i < lines.length; i++) {
            var obj = {};
            var currentline = lines[i];
            for (var j = 0; j < headers.length; j++) {
                obj[headers[j]] = currentline[j];
            }
            result.push(obj);
        }
        return result; //JSON
    }

    const fileUploaded = (data: any) => {
        let object = csvJSON(data);
        object.map(user => {
            props.inviteUser(user.email, user.redirect_url)
            .then((res) => {
                props.updateUser(res.invitedUser.id, user)
                .then(() => {
                    shareSiteWithUser(user.email, user.redirect_url)
                    .then(()=>{
                        console.log('Successful')
                    })
                })
            })
        })
    } 
    return (
        <>
            <div>Bulk Invite Component</div>
            <CSVReader onFileLoaded={(data, fileInfo) => { fileUploaded(data) }} />
        </>
    )
}

export default BulkInvite;