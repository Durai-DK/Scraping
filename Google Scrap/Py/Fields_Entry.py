import form

driver_path = r"D:\Durai\Driver\chromedriver.exe"
Excel_Sheet_save_location = r"D:\Durai\Scraping\Google Scrap\Save Data's\ "


gm = form.GoogleMaps(Title="Music composers",Location="Dubai",Save_Excel=Excel_Sheet_save_location)
gm.search_text()
