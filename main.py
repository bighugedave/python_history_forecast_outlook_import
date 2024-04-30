from import_history_forecast import ImportHistoryForecast

hcode = '1216'
save_folder_path = 'C:\\Users\\dayresx\\Downloads\\pulse\\{}'
target_subject = '1216 history_forecast'
import_hf = ImportHistoryForecast(xml_dir=save_folder_path, subject_text=target_subject)
import_hf.get_attachments()

