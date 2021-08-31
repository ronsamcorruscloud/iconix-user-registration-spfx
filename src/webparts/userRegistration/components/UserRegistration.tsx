import * as React from 'react';
import styles from './UserRegistration.module.scss';
import { IUserRegistrationProps } from './IUserRegistrationProps';
import { Label, Pivot, PivotItem } from '@fluentui/react';
import { Formik, Field, Form, FormikHelpers } from 'formik';
import { MessageBar, MessageBarType, Spinner } from '@fluentui/react';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { DefaultButton, PrimaryButton, MessageBarButton, IButtonStyles } from '@fluentui/react/lib/Button';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';

import { MSGraphClient } from '@microsoft/sp-http';
import { IUserRegistrationState } from './IUserRegistrationState';
import * as Yup from 'yup';
import { getSiteGroups, shareSiteWithUser, updateUserByID } from '../requests';
import BulkInvite from './bulk-invite';

const UserRegistrationSchema = Yup.object().shape({
  firstName: Yup.string()
    .min(2, 'Too Short!')
    .max(50, 'Too Long!')
    .required('Required'),
  lastName: Yup.string()
    .min(2, 'Too Short!')
    .max(50, 'Too Long!')
    .required('Required'),
  email: Yup.string().email('Invalid email').required('Required'),
  company: Yup.string()
    .min(2, 'Too Short!')
    .max(50, 'Too Long!')
    .required('Required'),
  phone: Yup.number(),
  address: Yup.string(),
  job_title: Yup.string().required('Required')
});

const narrowTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300, marginBottom: 10 } };
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300, marginBottom: 10 } };
const buttonStyles: IButtonStyles = { root: { marginTop: 20, marginBottom: 20 } }

const urls = [
  {
    key: 'https://iconixbrandgroup.sharepoint.com/sites/IconixHub',
    text: 'Iconix Hub'
  },
  {
    key: 'https://iconixbrandgroup.sharepoint.com/sites/LeeCooperHub',
    text: 'Lee Cooper Hub'
  },
  {
    key: 'https://iconixbrandgroup.sharepoint.com/sites/UmbroHub',
    text: 'Umbro Hub'
  }
]

export default class UserRegistration extends React.Component<IUserRegistrationProps, IUserRegistrationState> {
  constructor(props) {
    super(props);
    this.state = {
      success: null,
      message: '',
      options: [],
      tempID: null,
      isLoading: false
    }
  }
  public componentWillMount = () => {
    getSiteGroups().then(res => {
      console.log('groups', res);
      let options = []
      res.map(item => {
        options.push({ key: item.Id, text: item.Title })
      })
      this.setState({ options })
    })
  }
  public inviteUser = (email, url): Promise<any[]> => {
    return (
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): Promise<any[]> => {
          const invitation = {
            invitedUserEmailAddress: email,
            inviteRedirectUrl: url,
            sendInvitationMessage: true
          };
          return (
            client.api('/invitations')
              .post(invitation)
              .then(res => {
                console.log(res)
                return res
              })
              .catch(err => {
                throw err
              })
          )
        })
    )
  }
  public updateUser = (userID, data): Promise<any[]> => {
    return (
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): Promise<any[]> => {
          const user = {
            displayName: data.firstName + ' ' + data.lastName,
            mail: data.email,
            companyName: data.company,
            jobTitle: data.job_title
          };
          if (data.phone) {
            user['mobilePhone'] = data.phone;
          }
          if (data.address) {
            user['streetAddress'] = data.address;
          }
          return (
            client.api(`/users/${userID}`)
              .update(user)
              .then((res) => {
                console.log(res)
                return res
              })
              .catch(err => {
                throw err
              })
          )
        })
    )
  }
  public render(): React.ReactElement<IUserRegistrationProps> {
    return (
      <div className={styles.userRegistration}>
        {this.state.isLoading &&
          <div className={styles.loader}>
            <Spinner label="Inviting the user" />
          </div>
        }
        <div className={styles.container}>
          {this.state.message &&
            <MessageBar
              actions={
                <div>
                  <MessageBarButton onClick={() => this.setState({ success: null, message: '' })}>Cancel</MessageBarButton>
                </div>
              }
              messageBarType={this.state.success ? MessageBarType.success : MessageBarType.error}
              isMultiline={false}
            >
              {this.state.message}
            </MessageBar>
          }
          <Pivot aria-label="Large Link Size Pivot Example" linkSize="large">
            <PivotItem headerText="My Files">
              <>
                <h1 className={styles.formHeader}>Invite a User</h1>
                <Formik
                  initialValues={{
                    firstName: '',
                    lastName: '',
                    email: '',
                    company: '',
                    phone: '',
                    address: '',
                    job_title: '',
                    group: '',
                    url: ''
                  }}
                  validationSchema={UserRegistrationSchema}
                  onSubmit={(values, { resetForm }) => {
                    this.setState({ success: null, message: '', isLoading: true })
                    this.inviteUser(values.email, values.url['key'])
                      .then((res: any) => {
                        updateUserByID(res.invitedUser.id, values, this.props.context)
                          // this.updateUser(res.invitedUser.id, values)
                          .then(response => {
                            shareSiteWithUser(values.email, values.group['key'])
                              .then(result => {
                                this.setState({ success: true, message: "Successfully invited the user", isLoading: false });
                                resetForm()
                              })
                              .catch(err => {
                                this.setState({ success: false, message: "Could not add user to the sharepoint group", isLoading: false })
                              })
                          })
                          .catch(err => {
                            this.setState({ success: false, message: "Error happened during updating details. Try submitting again", isLoading: false })
                          })
                      })
                      .catch(err => {
                        this.setState({ success: false, message: "Error inviting the user", isLoading: false })
                      })
                  }}
                >
                  {({
                    values,
                    errors,
                    touched,
                    handleChange,
                    handleBlur,
                    handleSubmit,
                    resetForm,
                    isSubmitting,
                    setFieldValue
                    /* and other goodies */
                  }) => {
                    return (
                      <div>
                        <TextField
                          label={"First Name"}
                          value={values.firstName}
                          onChange={handleChange('firstName')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          required
                          errorMessage={errors.firstName && touched.firstName ? errors.firstName : null}
                        />
                        <TextField
                          label={"Last Name"}
                          value={values.lastName}
                          onChange={handleChange('lastName')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          required
                          errorMessage={errors.lastName && touched.lastName ? errors.lastName : null}
                        />
                        <TextField
                          label={"Email"}
                          value={values.email}
                          onChange={handleChange('email')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          required
                          errorMessage={errors.email && touched.email ? errors.email : null}
                        />
                        <TextField
                          label={"Company"}
                          value={values.company}
                          onChange={handleChange('company')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          required
                          errorMessage={errors.company && touched.company ? errors.company : null}
                        />
                        <TextField
                          label={"Job Title"}
                          value={values.job_title}
                          onChange={handleChange('job_title')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          required
                          errorMessage={errors.job_title && touched.job_title ? errors.job_title : null}
                        />
                        <TextField
                          label={"Phone"}
                          value={values.phone}
                          onChange={handleChange('phone')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          errorMessage={errors.phone && touched.phone ? errors.phone : null}
                        />
                        <TextField
                          label={"Address"}
                          value={values.address}
                          onChange={handleChange('address')}
                          styles={narrowTextFieldStyles}
                          onBlur={handleBlur}
                          errorMessage={errors.address && touched.address ? errors.address : null}
                        />
                        <Dropdown
                          label="Group"
                          selectedKey={values.group ? values.group['key'] : undefined}
                          // eslint-disable-next-line react/jsx-no-bind
                          onChange={(e, value) => {
                            setFieldValue('group', value)
                          }}
                          placeholder="Select an option"
                          options={this.state.options}
                          styles={dropdownStyles}
                        />
                        <Dropdown
                          label="Redirect URL"
                          selectedKey={values.url ? values.url['key'] : undefined}
                          // eslint-disable-next-line react/jsx-no-bind
                          onChange={(e, value) => {
                            setFieldValue('url', value)
                          }}
                          placeholder="Select an option"
                          options={urls}
                          styles={dropdownStyles}
                        />
                        <PrimaryButton onClick={() => handleSubmit()} styles={buttonStyles}>Submit</PrimaryButton>
                      </div>
                    )
                  }}

                </Formik>
              </>
            </PivotItem>
            <PivotItem headerText={"Bulk Invite"}>
                  <BulkInvite
                    inviteUser={this.inviteUser}
                    updateUser={(id, values) => updateUserByID(id, values, this.props.context)}
                    
                  />
            </PivotItem>
          </Pivot>
        </div>
      </div>
    );
  }
}
