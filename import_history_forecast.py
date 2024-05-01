import win32com.client as client


class ImportHistoryForecast:
    def __init__(self, xml_dir, subject_text):
        self.save_folder_path = xml_dir
        self.target_subject = subject_text

    def get_attachments(self):
        try:
            # create instance of Outlook
            outlook_app = client.Dispatch('Outlook.Application')

            # get the inbox
            namespace = outlook_app.GetNamespace('MAPI')
            inbox = namespace.GetDefaultFolder(6)

            # get only mail items from the inbox (other items can exists and will return an error if you try get the subject line of a non-mail item)
            mail_items = [item for item in inbox.Items if item.Class == 43 and self.target_subject in item.Subject]

            # filter to the target email
            filtered = [item for item in mail_items if self.target_subject in item.Subject]

            # get the first item if it exists (assuming the there is only one item to get)
            if len(mail_items) != 0:
                for mail_item in mail_items:
                    # get attachments
                    if mail_item.Unread and mail_item.Attachments.Count > 0:
                        attachments = mail_item.Attachments
                        mail_item.Unread = True
                        mail_item.Save()

                        # save the attachments
                        for file in attachments:
                            file.SaveAsFile('{}\\{}'.format(self.save_folder_path, file.FileName))

                outlook_app.Quit()
        except AttributeError as e:
            print(e)
