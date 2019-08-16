from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from pandas import ExcelWriter
import requests
import pandas as pd
import numpy as np
import time


df = pd.read_excel(r'C:\Users\dgreally\Downloads\ExcelDatabase.xlsx')
"""journals = df.loc[(df['Portfolio'] == 'STMS Physical Sciences & Engineering'), 'Journal Name'].drop_duplicates()"""

dataframes = pd.DataFrame()
journals = ['research synthesis methods']
# Opens the browser window and retrieves the initial url
browser = webdriver.Chrome(r'C:\Users\dgreally\Downloads\chromedriver_win32\chromedriver')
browser.maximize_window()
browser.get('https://error.incites.clarivate.com/error/Error?DestApp=IC2JCR&Error=IPError&Params=DestApp%3DIC2JCR&Route'
            'rURL=https%3A%2F%2Flogin.incites.clarivate.com%2F&Domain=.clarivate.com&Src=IP&Alias=IC2'
            )
browser.implicitly_wait(20)

# Selects Wiley from drop-down list and presses the 'Go' button
first_button = Select(browser.find_element_by_name('select2'))
first_button.select_by_visible_text('Wiley')
second_button = browser.find_element_by_class_name('shib_go').click()
browser.implicitly_wait(30)

for journal in journals:
    skip = 0
    journal = journal.strip()
    names = []
    years = []
    total_cites = []
    journal_impact_factors = []
    impact_factors_without_journal_self_cites = []
    five_year_impact_factors = []
    immediacy_indices = []
    citable_items = []
    cited_half_lives = []
    citing_half_life = []
    eigenfactor_scores = []
    article_influence_scores = []
    percentage_articles_in_citable_items = []
    normalised_eigenfactors = []
    average_JIF_percentiles = []

# Types journal name into search bar
    while True:
        searchbar = browser.find_element_by_id('search-inputEl')
        searchbar.clear()
        searchbar.send_keys(journal)
        time.sleep(5)
        try:
            browser.find_element_by_class_name('x-boundlist-item').click()
            time.sleep(5)
            break
        except:
            skip = 1
            break

    if skip == 1:
        print('Could not receive data for ' + journal + '.  Check spelling!')
        continue

    windows = browser.window_handles
    browser.switch_to.window(windows[1])
    browser.find_element_by_class_name('all-years-tab').click()
    time.sleep(10)
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    table = soup.select('[data-boundview=gridview-1011]')

    for column in table:
        year = column.select('.x-grid-cell-gridcolumn-1010 .x-grid-cell-inner a')[0].text
        years.append(year)

        total_cite = column.select('.x-grid-cell-Totalcites div a')[0].text
        total_cites.append(total_cite)

        journal_impact_factor = column.select('.x-grid-cell-JournalImpactFactor div a')[0].text
        journal_impact_factors.append(journal_impact_factor)

        impact_factor_without_journal_self_cites = column.select('.x-grid-cell-ImpactFactorWithoutJou'
                                                                 'rnalSelfCites div a')[0].text
        impact_factors_without_journal_self_cites.append(impact_factor_without_journal_self_cites)

        five_year_impact_factor = column.select('.x-grid-cell-FiveYearImpactFactor div a')[0].text
        five_year_impact_factors.append(five_year_impact_factor)

        immediacy_index = column.select('.x-grid-cell-ImmediacyIndex div a')[0].text
        immediacy_indices.append(immediacy_index)

        citable_item = column.select('.x-grid-cell-CitableItems div a')[0].text
        citable_items.append(citable_item)

        cited_half_life = column.select('.x-grid-cell-CitedHalfLife div a')[0].text
        cited_half_lives.append(cited_half_life)

        eigenfactor_score = column.select('.x-grid-cell-EigenFactorScore div a')[0].text
        eigenfactor_scores.append(eigenfactor_score)

        article_influence_score = column.select('.x-grid-cell-ArticleInfluenceScore div a')[0].text
        article_influence_scores.append(article_influence_score)

        percentage_article_in_citable_items = column.select('.x-grid-cell-OriginalResearch')[0].text
        percentage_articles_in_citable_items.append(percentage_article_in_citable_items)

        normalised_eigenfactor = column.select('.x-grid-cell-NormEigenFactor div a')[0].text
        normalised_eigenfactors.append(normalised_eigenfactor)

        average_JIF_percentile = column.select('.x-grid-cell-AvgJifPercentile div a')[0].text
        average_JIF_percentiles.append(average_JIF_percentile)

        names.append(journal)

    journal_data = pd.DataFrame({'name': names,
                                 'year': years,
                                 'total cites': total_cites,
                                 'journal impact factor': journal_impact_factors,
                                 'impact factor without journal self cites': impact_factors_without_journal_self_cites,
                                 '5 year impact factor': five_year_impact_factors,
                                 'immediacy index': immediacy_indices,
                                 'citable items': citable_items,
                                 'cited half-life': cited_half_lives,
                                 'eigenfactor score': eigenfactor_scores,
                                 'article influence score': article_influence_scores,
                                '% articles in citable items': percentage_articles_in_citable_items,
                                 'normalised eigenfactor': normalised_eigenfactors,
                                 'average JIF percentile': average_JIF_percentiles
                                 })

    dataframes = dataframes.append(journal_data, ignore_index=True)
    browser.close()
    browser.implicitly_wait(2)
    browser.switch_to.window(browser.window_handles[0])
    searchbar.clear()

browser.close()

while True:
    excel = input('Note: Rename previous export or it will be overwritten!\nExport to CSV? (Y/N) ')
    if excel == 'Y':
        dataframes.to_csv(r'C:\Users\dgreally\OneDrive - Wiley\Documents\dataframe.csv')
        break
    elif excel == 'N':
        print('goodbye...')
        break
    else:
        print('try again...')