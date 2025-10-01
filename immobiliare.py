from bs4 import BeautifulSoup
import requests
import pandas as pd
import re
from  selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from geopy.geocoders import Nominatim
from excel_manipulation import correct_excel, generate_definitive_table
from openpyxl import load_workbook
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from requests.adapters import HTTPAdapter
import socket
from urllib3.connection import HTTPConnection
import geo_data as gd

class ReuseConnectionAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        kwargs['socket_options'] = [(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)]
        return super(ReuseConnectionAdapter, self).init_poolmanager(*args, **kwargs)

def setup_minimized_browser():
    try:
        # Set up Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-background-timer-throttling")
        chrome_options.add_argument("--disable-backgrounding-occluded-windows")
        chrome_options.add_argument("--disable-renderer-backgrounding")
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument("window-size=50,50")  # Start with a very small window

        # Suppress logging
        chrome_options.add_argument("--log-level=3")
        
        # Use the cached ChromeDriver path
        driver_path = ChromeDriverManager().install()
        correct_driver_path = os.path.join(os.path.dirname(driver_path), 'chromedriver.exe')
        
        if not os.path.exists(correct_driver_path):
            raise FileNotFoundError(f"ChromeDriver executable not found at {correct_driver_path}")
        
        driver = webdriver.Chrome(service=Service(correct_driver_path), options=chrome_options)

        # Minimize the browser window
        driver.minimize_window()
        
        return driver
    
    except Exception as e:
        print(f"An error occurred while setting up the browser: {e}")
        raise



def parse_features(dict_features, results):
    all_types = ['#planimetry', '#size', '#bath', '#stairs', '#elevator', '#balcony', '#beach-umbrella', '#couch-lamp']
    for result in results:
        list = result.find_all('div', {'class':"in-listingCardFeatureList__item"})
        types = [l.use['xlink:href'] for l in list]
        values = [l.get_text() for l in list]
        for type in all_types:
            if type in types:
                dict_features[type].append(values[types.index(type)])
            else:
                dict_features[type].append('null')
    return dict_features

def get_immobiliare(area,city):
    print(f"Processing {area}")

    address = []
    results = []
    prices = []
    links = []
    area_info = []
    exact_address = []
    location=[]
    energy_class = []
    energy_values = []
    all_types = ['#planimetry', '#size', '#bath', '#stairs', '#elevator', '#balcony', '#beach-umbrella', '#couch-lamp']
    features = {type:[] for type in all_types}
    get_coordinates = False
    browser = setup_minimized_browser()

    # Estrae le informazioni dai badge di pagina
    website = f'https://www.immobiliare.it/vendita-case/{area}'
    headers = ({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62'})
    session = requests.Session()
    adapter = ReuseConnectionAdapter()
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    response = session.get(website, headers=headers)
    html = browser.page_source
    soup = BeautifulSoup(response.content, 'html.parser')

    breadcrumb_links = soup.find_all('a', {'class': 'in-breadcrumbLink__dropdownLink'})
    breadcrumb_links = [l['href'] for l in breadcrumb_links]
    if breadcrumb_links == []:
        breadcrumb_links.append(f'https://www.immobiliare.it/vendita-case/{area}')

    j = 0
    for l in breadcrumb_links:
        i = 1
        j += 1
        valid_page = True
        print(f'Scanning area {j}/{len(breadcrumb_links)}')
        while(valid_page):
            website = f'{l}/?criterio=rilevanza&pag={i}'
            headers = ({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62'})
            session = requests.Session()
            adapter = ReuseConnectionAdapter()
            session.mount('http://', adapter)
            session.mount('https://', adapter)
            response = session.get(website, headers=headers)
            if response.status_code == 404:
                valid_page = False
            else:
                soup = BeautifulSoup(response.content, 'html.parser')
                results = soup.find_all('div', {'class' : 'nd-mediaObject__content in-listingCardPropertyContent'})
                
                features= parse_features(features, results)
                address +=  [result.find('a', {'class':"in-listingCardTitle"}).get_text() for result in results]
                links += [result.find('a', {'class':"in-listingCardTitle"})['href'] for result in results]
                prices += [result.find('div', {'class': 'in-listingCardPrice'}).get_text() for result in results]
                real_estate = pd.DataFrame(columns=['Link','Indirizzo','Civico','Citta-zona-via-civico', 'Prezzo', 'n locali', 'area','n bagni','piano','ascensore?','balcone?','terrazzo?','arredato?','Classe energetica','Efficienza energetica (numero)'])
                i+=1
    d = {'#':0}
    pattern = r'\b\d{2}/\d{2}/\d{4}\b'
    

    #Entra in ciascun link
    for link in links:
        not_grouped_titles = []
        not_grouped_attr = []
        headers = ({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62'})
        browser.get(link)

        try:
            button = browser.find_element(By.XPATH, "//button[@class='nd-button re-primaryFeatures__openDialogButton']")
            # If the button is found, click it
            browser.execute_script("arguments[0].click();", button)
        except NoSuchElementException:
        # Button is not present, continue with the alternative content extraction
            pass

        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser')

        titles = soup.find_all('dt', {'class':'re-primaryFeaturesDialogSection__featureTitle'})
        attr = soup.find_all('dd', {'class':"re-primaryFeaturesDialogSection__featureDescription"})

        # If the button-specific content is not found, fall back to the alternative
        if not titles or not attr:
            titles = soup.find_all('dt', {'class': 're-featuresItem__title'})
            attr = soup.find_all('dd', {'class': "re-featuresItem__description"})
        
        # Step 2: Close the modal/tab by clicking the 'X' button
        try:
            close_button = browser.find_element(By.XPATH, "//button[@class='nd-dialogFrame__close']")  # Update with correct XPath
            browser.execute_script("arguments[0].click();", close_button)
        except NoSuchElementException:
            pass

        # Step 3: Click the 'Vedi dettaglio' button
        try:
            vedi_dettaglio_button = browser.find_element(By.XPATH, "//button[@class='nd-button re-energy__openButton']")
            browser.execute_script("arguments[0].click();", vedi_dettaglio_button)
        except NoSuchElementException:
            pass

        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser')

        # Locate the 'Invernale' and 'Estivo' under 'Prestazione energetica del fabbricato'
        prestazione_section = soup.find('dd', {'class': 're-energy__featureSeasons'})
        if prestazione_section:
            # Locate the 'Invernale' SVG and check the xlink:href value
            invernale_icon = prestazione_section.find_next('span').find('use')
            invernale_value = 'Si' if invernale_icon and invernale_icon['xlink:href'] == '#face-high' else 'No'

            # Locate the 'Estivo' SVG and check the xlink:href value
            estivo_span = prestazione_section.find_next('span').find_next_sibling('span')
            if estivo_span:
                estivo_icon = estivo_span.find('use')
                estivo_value = 'Si' if estivo_icon and estivo_icon['xlink:href'] == '#face-high' else 'No'
            else:
                estivo_value = 'No'
        else:
            invernale_value = 'No'
            estivo_value = 'No'

        not_grouped_titles += ['Prestazione energetica estiva?', 'Prestazione energetica invernale?']
        not_grouped_attr += [estivo_value, invernale_value]

        # Locate the 'Indice prest. energetica rinnovabile'
        indice_prest_section = soup.find('dt', text='Indice prest. energetica rinnovabile')
        if indice_prest_section:
            indice_value = indice_prest_section.find_next_sibling('dd').text.strip()
        else:
            indice_value = 'null'
        
        not_grouped_titles.append('Indice prest. energetica rinnovabile')
        not_grouped_attr.append(indice_value)


        others_div = soup.find('ul',{'class': 're-featuresBadges__list'})
        others = [t.get_text() for t in others_div.find_all('div',{'class': 'nd-badge'})] if others_div is not None else []

        spese_label = soup.find('dt', text='spese condominio')
        if spese_label:
            # Get the corresponding value
            spese_value = spese_label.find_next_sibling('dd').text.strip()
        else:
            spese_value = 'null'
        
        not_grouped_titles.append('Spese condominio')
        not_grouped_attr.append(spese_value)

        area_info += [tuple([el.get_text() for el in soup.find_all('span',{'class':'re-title__location'})])]
        desc_div = [b.get_attribute('textContent').strip() for b in browser.find_elements(By.CLASS_NAME, 're-locationInfo')]
            
        ec = soup.find('span',{'class':"re-mainConsumptions__energy"})
        energy_class += ['null' if ec is None else ec.get_text()]

        ev = soup.find('div',{'class':'re-mainConsumptions'})
        if ev is not None and len(ev.find_all("p")) > 1:
            if ev.find_all("p")[1].get_text().split(",")[0] != '':
                energy_values += [ev.find_all("p")[1].get_text().split(",")[0]]
            else:
                energy_values += ['null']
        else:
            energy_values += ['null']

        if desc_div != []:
            exact_address += [desc_div[-1]]
        else:
            exact_address += ['null']
            
        scripts = soup.find_all('script')
        coordinates = None
        for script in scripts:
            if 'latitude' in script.text and 'longitude' in script.text:
                coordinates = script.text
                break
            
            
        if coordinates:
            # Extract latitude and longitude from the script text
            start_lat = coordinates.find('latitude') + 10
            end_lat = coordinates.find(',', start_lat)
            latitude = coordinates[start_lat: end_lat].strip()

            start_lon = coordinates.find('longitude') + 11
            end_lon = coordinates.find(',', start_lon)
            longitude = coordinates[start_lon: end_lon].strip()
                
            location.append((latitude, longitude))
        else:
            location.append(('null','null'))

        string_others = ""
        for i in range(len(others)):
            string_others += others[i] + ';'
        if string_others == "":
            string_others = 'null'

        not_grouped_titles.append('altre caratteristiche')
        not_grouped_attr.append(string_others)

        titles = [t.get_text() for t in titles] + not_grouped_titles
        attr = [a.get_text() for a in attr] + not_grouped_attr
        d['#'] += 1
        print(f"Processing article n°{d['#']} out of {len(links)}")

        for i in range(len(titles)):
            if titles[i] == 'Riferimento e Data annuncio':
                attr[i] = re.findall(pattern,attr[i])[0]
            if titles[i] not in d:
                if d['#'] == 1:
                    d[titles[i]] = [attr[i]]
                else:
                    n_null = d['#'] - 1
                    d[titles[i]] = ['null' for _ in range(n_null)]
                    d[titles[i]].append(attr[i])

            else:
                if len(d[titles[i]]) < d['#']:
                    d[titles[i]].append(attr[i])

        for k in d:
            if k != '#' and len(d[k]) < d['#']:
                d[k].append('null')


    area_info = [f'{t[0]};{t[1]};{t[2]}' if len(t) >= 3 else f'{t[0]};{t[1]}' if len(t) == 2 else f'{t[0]}' if len(t) == 1 else 'null' for t in area_info]
    for i in range (len(links)):
        real_estate=real_estate._append({'Link':links[i],'Indirizzo':exact_address[i], 'Coord':location[i], 
                                         'Citta-zona-via-civico':area_info[i], 'Prezzo':prices[i], 'n locali':features['#planimetry'][i], 'area':features['#size'][i],
                                        'n bagni':features['#bath'][i], 'piano':features['#stairs'][i], 'ascensore?':features['#elevator'][i],
                                        'balcone?':features['#balcony'][i], 'terrazzo?':features['#beach-umbrella'][i], 'arredato?':features['#couch-lamp'][i], 'Classe energetica':energy_class[i], 'Efficienza energetica (numero)':energy_values[i]}, ignore_index=True)

    for k in d:
        if k != '#':
            real_estate.insert(len(real_estate.columns), k, d[k], True)
    area = area.replace('/','_')
    #to Excel
    real_estate.to_excel(f'{city}/real estate_{area}.xlsx')
    session.close()

if __name__ == '__main__':
    places = ['Cagliari','Roma', 'Milano']
    for city in places:
        if not os.path.exists(city):
            os.makedirs(city)
        comuni = gd.get_comuni(gd.get_sigla(city))
        for place in comuni:
            place = place.replace("'","-")
            if len(place.split()) > 1:
                place = place.replace(" ","-")
            if not os.path.exists(f'{city}/real estate_{place}.xlsx') or not os.path.exists(f'{city}/{place}.xlsx') or not os.path.exists(f'{city}/{place}-Def.xlsx'):
                get_immobiliare(place, city)
                wb = load_workbook(filename=f'{city}/real estate_{place}.xlsx')
                sheet = wb.active
                wb.save(f'{city}/{place}.xlsx')
                correct_excel(f'{city}/{place}.xlsx')
                generate_definitive_table(f'{city}/{place}.xlsx',f'{city}/{place}-Def.xlsx')
                print(f'{place} in provincia di {city} completato!')
                