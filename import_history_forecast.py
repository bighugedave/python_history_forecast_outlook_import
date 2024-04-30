import win32com.client as client


class ImportHistoryForecast:
    def __init__(self, xml_dir, subject_text):
        self.save_folder_path = xml_dir
        self.target_subject = subject_text

    def get_attachments(self):
        # create instance of Outlook
        outlook = client.Dispatch('Outlook.Application')

        # get the inbox
        namespace = outlook.GetNameSpace('MAPI')
        inbox = namespace.GetDefaultFolder(6)

        # get only mail items from the inbox (other items can exists and will return an error if you try get the subject line of a non-mail item)
        mail_items = [item for item in inbox.Items if item.Class == 43]

        # filter to the target email
        filtered = [item for item in mail_items if self.target_subject in item.Subject]

        # get the first item if it exists (assuming the there is only one item to get)
        if len(filtered) != 0:
            target_email = filtered[0]
            # get attachments
            if target_email.Unread and target_email.Attachments.Count > 0:
                attachments = target_email.Attachments
                target_email.Unread = False
                target_email.Save()

                # save the attachments
                for file in attachments:
                    file.SaveAsFile(self.save_folder_path.format(file.FileName))
