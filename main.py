from selenium.common.exceptions import StaleElementReferenceException
from utils import *

PUBLI_CLASS = "div[class='x78zum5 xdt5ytf x5yr21d xa1mljc xh8yej3 x1bs97v6 x1q0q8m5 xso031l x11aubdm xnc8uc2']"
LINK_ANUNCIO = "a[class='_ad63']"
CURTIDAS = "span[class='html-span xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x1hl2dhg x16tdsg8 x1vvkbs']"
SCROLL_AMOUNT = 1500
NOME_PLANILHA_EXCEL = "urls_dropshipping.xlsx"
QUANTIDADE_LIKES_PUBLICACAO = 100


def main():
    driver = start_navegador()
    divs_publicacao = driver.find_elements(By.CSS_SELECTOR, PUBLI_CLASS)

    while True:
        try:
            for publi in divs_publicacao:
                if publi is not None:
                    if "Patrocinado" in publi.text:
                        likes = publi.find_element(By.CSS_SELECTOR, CURTIDAS)
                        likes_int = int(likes.text.replace(".", ""))

                        if likes_int > QUANTIDADE_LIKES_PUBLICACAO:
                            try:
                                ad = publi.find_element(By.CSS_SELECTOR, LINK_ANUNCIO)
                                ad.click()
                                fechar_aba_anuncio(driver)
                                link_ad = ad.get_attribute("href")

                                salvar_excel(
                                    link_ad,
                                    likes_int,
                                    publi.text.replace("\n", "-"),
                                    NOME_PLANILHA_EXCEL,
                                )
                            except:
                                break
        except StaleElementReferenceException:
            divs_publicacao = driver.find_elements(By.CSS_SELECTOR, PUBLI_CLASS)

        divs_publicacao = scrollar(driver)


if __name__ == "__main__":
    main()
