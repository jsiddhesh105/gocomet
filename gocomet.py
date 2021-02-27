from selenium import webdriver
import xlsxwriter
from webdriver_manager.chrome import ChromeDriverManager
driver = webdriver.Chrome(ChromeDriverManager().install())

all_product_amazon = []
all_product_flipkart = []

def get_product_details_amazon(link):
    driver.get(link)
    source = "amazon"
    try:
        title = driver.find_element_by_id('productTitle').text
    except:
        title = "Not available"
    
    
    try:
        price = driver.find_element_by_xpath('//*[@id="priceblock_ourprice"]').text
    except:
        price = "Not available"
        
    try:
        delivery_time = driver.find_element_by_xpath('//*[@id="ddmDeliveryMessage"]/b').text
    except:
        delivery_time = "Not available"
        
    try:
        model_number = driver.find_element_by_xpath('//span[contains(text(),"Model")]/following-sibling::span').text
    except:
        try:
            model_number = driver.find_element_by_xpath('//th[contains(text(),"Item model number")]/following-sibling::td').text
        except:
            model_number = "Not available"
            
    return title,model_number,source,price,delivery_time,link
    
    


def product_from_amazon(prodcut):
    
    driver.get('https://www.amazon.in')
    
    search_box = driver.find_element_by_id('twotabsearchtextbox').send_keys(product)
    search_button = driver.find_element_by_id("nav-search-submit-button").click()
    
    driver.implicitly_wait(3)
    
    n = 0
    c = 0
    product_link = []
    while c<10:
            l= '//span[@data-component-type="s-product-image" and @data-component-id="'+str(n)+'"]'
            print(l)
            try:
                product_link.append(driver.find_element_by_xpath(l+'/a').get_attribute('href'))
                n = n + 1
                c = c + 1
            except:
                n = n + 1
                continue
            
    for link in product_link:
        details = get_product_details_amazon(link)
        all_product_amazon.append(details)
        
    return all_product_amazon
        
def get_product_details_flipkart(link):
    driver.get(link)
    driver.implicitly_wait(3)
    source = "flipkart"
    try:
        title = driver.find_element_by_class_name('yhB1nd').text
    except:
        try:
            title = driver.find_element_by_class_name('_4rR01T').text
        except:
            title = "Sorry Could not Find"
    
    try:
        price = driver.find_element_by_class_name('_30jeq3').text
    except:
        price = "Not Available"
        
    try:
        delivery_time = driver.find_element_by_class_name('_1TPvTK').text
    except:
        delivery_time = "Not Available"
        
    try:
        model_number = driver.find_element_by_xpath('//td[contains(text(),"Model Name")]/following-sibling::td').text
    except:
        try:
            model_number = driver.find_element_by_xpath('//td[contains(text(),"Model Number")]/following-sibling::td').text
        except:
            model_number = "Not available"
            
    return title,model_number,source,price,delivery_time,link
    
    
    
        
def product_from_flipkar(product):
#     driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get('https://www.flipkart.com')
    driver.implicitly_wait(3)
    
    search_box = driver.find_element_by_class_name('_3704LK').send_keys(product)
    driver.find_element_by_xpath("/html/body/div[2]/div/div/button").click()
    driver.find_element_by_xpath("//*[@id='container']/div/div[1]/div[1]/div[2]/div[2]/form/div/button").click()
    
    driver.implicitly_wait(3)
    
    n = 2
    c = 0
    product_link = []
    while c<10:
            l= '//*[@id="container"]/div/div[3]/div[1]/div[2]/div['+ str(n) +']/div/div/div/a'
            try:
                product_link.append(driver.find_element_by_xpath(l).get_attribute('href'))
                n = n + 1
                c = c + 1
            except:
                try:
                    product_link.append(driver.find_element_by_class_name('_2UzuFa').text)
                    c = c + 1
                    n = n + 1
                except:
                    continue



    for link in product_link:
        details = get_product_details_flipkart(link)
        all_product_flipkart.append(details)
        
    return all_product_flipkart
        

product = input("Enter the product:")

amazon_products = product_from_amazon(product)
flipkart_products = product_from_flipkar(product)
driver.close()

workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet() 
heading = ["Product Name","Model Number","Source","Price","Delivery_time","Link of the Product"]

print(amazon_products)

for i in range(len(heading)):
    worksheet.write(0, i, heading[i])
    
for i in range(0,len(amazon_products)):
    p = amazon_products[i]
    for j in range(len(p)):
        worksheet.write(i+1, j, p[j])

a = len(amazon_products)

for i in range(0,len(flipkart_products)):
    p = flipkart_products[i]
    for j in range(len(p)):
        worksheet.write(i+a+1, j, p[j])        

print("File generated Sucessfully...!")                
workbook.close()
   
