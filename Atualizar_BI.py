from selenium import webdriver
from selenium.webdriver.chrome.options import Options


options = Options()
options.add_experimental_option("detach", True)
browser = webdriver.Chrome(options=options)
browser.get("https://www.google.com.br")




