import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
from scipy.stats import ttest_ind
import xlsxwriter
import os, sys


def get_metal_stats(df, metal):
    stats = {}
    #grab data for given metal.
    metal_data = df[df['CHEMICAL_NAME'] == metal]
    #split into each unique loc_type
    loc_types = metal_data['LOC_TYPE'].unique()
    for loc_type in loc_types:
        #get data for each loc_type
        loc_type_data = metal_data[metal_data['LOC_TYPE'] == loc_type]
        print('')
        print(metal, loc_type, len(loc_type_data))

        stats[loc_type] = get_stats(loc_type_data)
    
    return stats

def get_stats(metal_data):
    stats = {'historic':{}, 'latest':{}}
    #split metal data between locations with 'REF' in the name and locations without 'REF' in the name
    metal_ref_data = metal_data[metal_data['LOC_ID'].str.contains('REF')]
    metal_site_data = metal_data[~metal_data['LOC_ID'].str.contains('REF')]
    #split into washed and unwashed data
    metal_ref_washed_data = metal_ref_data[metal_ref_data['LOC_SUBTYPE'] == 'WASHED']
    metal_ref_unwashed_data = metal_ref_data[metal_ref_data['LOC_SUBTYPE'] == 'UNWASH']
    metal_site_washed_data = metal_site_data[metal_site_data['LOC_SUBTYPE'] == 'WASHED']
    metal_site_unwashed_data = metal_site_data[metal_site_data['LOC_SUBTYPE'] == 'UNWASH']
    
    # stats['summary'] = {'ref': {'count':len(metal_ref_data.index), 'mean': metal_ref_data['REPORT_RESULT_VALUE'].mean(), 'std': metal_ref_data['REPORT_RESULT_VALUE'].std()}, 'site': {'count': len(metal_data.index), 'mean': metal_data['REPORT_RESULT_VALUE'].mean(), 'std': metal_data['REPORT_RESULT_VALUE'].std()}}
    #get the correlation over time
    stats['all_time'] = {'correlation':{}, 'mean':{}}
    stats['all_time']['correlation']['unwashed_ref'] = metal_ref_unwashed_data['REPORT_RESULT_VALUE'].corr(metal_ref_unwashed_data['SAMPLE_YEAR'])
    stats['all_time']['correlation']['unwashed_site'] = metal_site_unwashed_data['REPORT_RESULT_VALUE'].corr(metal_site_unwashed_data['SAMPLE_YEAR'])
    stats['all_time']['correlation']['washed_ref'] = metal_ref_washed_data['REPORT_RESULT_VALUE'].corr(metal_ref_washed_data['SAMPLE_YEAR'])
    stats['all_time']['correlation']['washed_site'] = metal_site_washed_data['REPORT_RESULT_VALUE'].corr(metal_site_washed_data['SAMPLE_YEAR'])

    #get the mean value
    stats['all_time']['mean']['unwashed_ref'] = metal_ref_unwashed_data['REPORT_RESULT_VALUE'].mean()
    stats['all_time']['mean']['unwashed_site'] = metal_site_unwashed_data['REPORT_RESULT_VALUE'].mean()
    stats['all_time']['mean']['washed_ref'] = metal_ref_washed_data['REPORT_RESULT_VALUE'].mean()
    stats['all_time']['mean']['washed_site'] = metal_site_washed_data['REPORT_RESULT_VALUE'].mean()

    
    #get stats for each year
    years = metal_data['SAMPLE_YEAR'].unique()
    for year in years:
        year_stats = {}
        year_data_ref = metal_ref_data[metal_ref_data['SAMPLE_YEAR'] == year]
        year_data_site = metal_site_data[metal_site_data['SAMPLE_YEAR'] == year]
        #split into washed and unwashed data
        year_data_ref_washed = year_data_ref[year_data_ref['LOC_SUBTYPE'] == 'WASHED']
        year_data_ref_unwashed = year_data_ref[year_data_ref['LOC_SUBTYPE'] == 'UNWASH']
        year_data_site_washed = year_data_site[year_data_site['LOC_SUBTYPE'] == 'WASHED']
        year_data_site_unwashed = year_data_site[year_data_site['LOC_SUBTYPE'] == 'UNWASH']

        #Get mean, std, count of each subset
        year_stats['washed_ref'] = {'count': len(year_data_ref_washed.index), 'mean': year_data_ref_washed['REPORT_RESULT_VALUE'].mean(), 'std': year_data_ref_washed['REPORT_RESULT_VALUE'].std()}
        year_stats['unwashed_ref'] = {'count': len(year_data_ref_unwashed.index), 'mean': year_data_ref_unwashed['REPORT_RESULT_VALUE'].mean(), 'std': year_data_ref_unwashed['REPORT_RESULT_VALUE'].std()}
        year_stats['washed_site'] = {'count': len(year_data_site_washed.index), 'mean': year_data_site_washed['REPORT_RESULT_VALUE'].mean(), 'std': year_data_site_washed['REPORT_RESULT_VALUE'].std()}
        year_stats['unwashed_site'] = {'count': len(year_data_site_unwashed.index), 'mean': year_data_site_unwashed['REPORT_RESULT_VALUE'].mean(), 'std': year_data_site_unwashed['REPORT_RESULT_VALUE'].std()}

        #get t-test of site vs ref for washed and unwashed
        t, p = ttest_ind(year_data_ref_washed['REPORT_RESULT_VALUE'], year_data_site_washed['REPORT_RESULT_VALUE'])
        year_stats['washed_ttest'] = {'t': t, 'p': p}
        t, p = ttest_ind(year_data_ref_unwashed['REPORT_RESULT_VALUE'], year_data_site_unwashed['REPORT_RESULT_VALUE']) 
        year_stats['unwashed_ttest'] = {'t': t, 'p': p}
        # print(year_stats)
        #is it significant?
        year_stats['washed_ttest']['significant'] = year_stats['washed_ttest']['p'] < 0.05
        year_stats['unwashed_ttest']['significant'] = year_stats['unwashed_ttest']['p'] < 0.05

        stats['historic'][year] = year_stats

        #check if this is the most recent year
        if year == max(years):
            stats['latest'] = year_stats
            stats['latest']['year'] = year

    #recursively through all nodes of stats dict and replace nan with None
    def replace_nan_with_none(d):
        for k, v in d.items():
            if isinstance(v, dict):
                replace_nan_with_none(v)
            elif isinstance(v, float) and np.isnan(v):
                d[k] = None
        return d

    return replace_nan_with_none(stats)

def scatter(metal_data, name=None, single=False):
    if not name: name = 'scatter'
    
    metal_ref_data = metal_data[metal_data['LOC_ID'].str.contains('REF')]
    metal_site_data = metal_data[~metal_data['LOC_ID'].str.contains('REF')]
    #does this data contain washed vs unwashed samples?
    if 'WASHED' in metal_data['LOC_SUBTYPE'].unique() and not single:
        #split both ref and non ref into washed and unwashed
        metal_ref_washed_data = metal_ref_data[metal_ref_data['LOC_SUBTYPE'] == 'WASHED']
        metal_ref_unwashed_data = metal_ref_data[metal_ref_data['LOC_SUBTYPE'].str.contains('UNWASH')]
        metal_site_washed_data = metal_site_data[metal_site_data['LOC_SUBTYPE'] == 'WASHED']
        metal_site_unwashed_data = metal_site_data[metal_site_data['LOC_SUBTYPE'].str.contains('UNWASH')]
        #graph data over time but use the same y axis
        plt.clf()
        plt.scatter(metal_ref_unwashed_data['SAMPLE_DATE'], metal_ref_unwashed_data['REPORT_RESULT_VALUE'], label='REF Unwashed', alpha=0.5)
        plt.scatter(metal_site_unwashed_data['SAMPLE_DATE'], metal_site_unwashed_data['REPORT_RESULT_VALUE'], label='Site Unwashed', alpha=0.5)
        #get the max value for the y axis
        max_y = metal_data['REPORT_RESULT_VALUE'].max()
        plt.ylim(0, max_y*1.1)
        plt.gca().xaxis.set_major_formatter(DateFormatter('%Y'))
        plt.legend()
        plt.title('Unwashed')
        plt.savefig(f'files/{name}_unwashed.png', bbox_inches='tight', dpi=300)
        # plt.show()

        plt.clf()
        plt.scatter(metal_ref_washed_data['SAMPLE_DATE'], metal_ref_washed_data['REPORT_RESULT_VALUE'], label='REF Washed', alpha=0.5)
        plt.scatter(metal_site_washed_data['SAMPLE_DATE'], metal_site_washed_data['REPORT_RESULT_VALUE'], label='Site Washed', alpha=0.5)
        plt.ylim(0, max_y*1.1)
        plt.gca().xaxis.set_major_formatter(DateFormatter('%Y'))
        plt.legend()
        plt.title('Washed')
        plt.savefig(f'files/{name}_washed.png', bbox_inches='tight', dpi=300)
        # plt.show()
    else:
        plt.clf()
        plt.scatter(metal_ref_data['SAMPLE_DATE'], metal_ref_data['REPORT_RESULT_VALUE'], label='REF', alpha=0.5)
        plt.scatter(metal_site_data['SAMPLE_DATE'], metal_site_data['REPORT_RESULT_VALUE'], label='Site', alpha=0.5)
        plt.legend()
        plt.gca().xaxis.set_major_formatter(DateFormatter('%Y'))
        # plt.show()
        #save plot image to files folder
        plt.savefig(f'files/{name}.png', bbox_inches='tight', dpi=300)

import numpy as np

def to_sheet(metal_data, metal_stats, workbook=None):
    current_path = os.path.dirname(os.path.abspath(__file__))
    if not workbook:
        filename =  os.path.join(current_path, "files/output.xlsx")
        workbook = xlsxwriter.Workbook(filename)
    metal_name = metal_data['CHEMICAL_NAME'].unique()[0]
    sheet = workbook.add_worksheet(metal_name.capitalize())
    
    #set first column extra wide
    sheet.set_column(0, 0, 32)
    #set all columns to 20 width
    sheet.set_column(1, 100, 18)
    #warning with a soft yellow background
    format_warning = workbook.add_format({'bg_color': '#f5cd6e'})
    format_fill = workbook.add_format({'bg_color': '#DCE6F1'})
    format_italic = workbook.add_format({'italic': True, 'font_size': 9})
    format_bold = workbook.add_format({'bold': True})
    #what is the largest number of decimals in the data?
    decimals = metal_data['REPORT_RESULT_VALUE'].apply(lambda x: len(str(x).split('.')[1]) if '.' in str(x) else 0).max()
    decimals = max(decimals, 2)
    number_format = f'#,##0.{decimals * "0"}'
    number_format = workbook.add_format({'num_format': number_format})
    workbook.add_format()
    #italic and size 9
    
    #Metal Name
    sheet.write(0, 0, metal_name.capitalize(), format_fill)
    #get the unit type (REPORT_RESULT_UNIT)
    unit_type = metal_data['REPORT_RESULT_UNIT'].unique()[0]

    sheet.write(1, 0, unit_type, format_italic)

    row, col = 2, 0
    for sample_type, stats in metal_stats.items():
        #Sample Type
        #get number of years in data
        years = metal_data['SAMPLE_YEAR'].unique()
        num_years = len(years)
        #Group/Merge Cells
        num_cols = max([num_years + 1, 15])
        sheet.merge_range(row, col, row, num_cols, sample_type.upper(), format_fill)
        row += 2

        #All Time
        sheet.write(row, col, 'All Time', format_bold)
        sheet.write_row(row, col+1, ['Unwashed Site', 'Unwashed Ref', 'Washed Site', 'Washed Ref'])

        row += 1
        sheet.write(row, col, 'Correlation with time (-1 to 1)')
        # write the correlation values and add warning format if above .3
        sheet.write(row, col+1, stats['all_time']['mean']['unwashed_site'], format_warning if abs(stats['all_time']['mean']['unwashed_site']) > .35 else None)
        sheet.write(row, col+2, stats['all_time']['mean']['unwashed_ref'], format_warning if abs(stats['all_time']['mean']['unwashed_ref']) > .35 else None)
        sheet.write(row, col+3, stats['all_time']['mean']['washed_site'], format_warning if abs(stats['all_time']['mean']['washed_site']) > .35 else None)
        sheet.write(row, col+4, stats['all_time']['mean']['washed_ref'], format_warning if abs(stats['all_time']['mean']['washed_ref']) > .35 else None)
        # sheet.write_row(row, col+1, [stats['all_time']['correlation']['unwashed_site'], stats['all_time']['correlation']['unwashed_ref'], stats['all_time']['correlation']['washed_site'], stats['all_time']['correlation']['washed_ref']], number_format)
        row += 1
        sheet.write(row, col, 'Mean value')
        sheet.write_row(row, col+1, [stats['all_time']['mean']['unwashed_site'], stats['all_time']['mean']['unwashed_ref'], stats['all_time']['mean']['washed_site'], stats['all_time']['mean']['washed_ref']], number_format)
        row += 2

        #Lasest Year
        latest_stats = stats['latest']
        sheet.write(row, col, f"Latest Year ({latest_stats['year']})", format_bold)
        sheet.write_row(row, col+1, ['p-value', 't-test'])
        row += 1

        format_good = workbook.add_format({'bg_color': '#C6EFCE'})
        format_bad = workbook.add_format({'bg_color': '#FFC7CE'})
        if not latest_stats['washed_ttest']['p']:
            sheet.write(row, col, 'Not enough Washed data for t test', format_good)
        else:
            if latest_stats['washed_ttest']['p'] < 0.05:
                sheet.write(row, col, 'Washed Site and Ref are different', format_bad)
            else:
                sheet.write(row, col, 'Washed Site and Ref are similar', format_good)
        sheet.write(row, col+1, latest_stats['washed_ttest']['p'], number_format)
        sheet.write(row, col+2, latest_stats['washed_ttest']['t'], number_format)
        
        row += 1
        if not latest_stats['unwashed_ttest']['p']:
            sheet.write(row, col, 'Not enough Unwashed data for t test', format_good)
        else:
            if latest_stats['unwashed_ttest']['p'] < 0.05:
                sheet.write(row, col, 'Unwashed Site and Ref are different', format_bad)
            else:
                sheet.write(row, col, 'Unwashed Site and Ref are similar', format_good)
        sheet.write(row, col+1, latest_stats['unwashed_ttest']['p'], number_format)
        sheet.write(row, col+2, latest_stats['unwashed_ttest']['t'], number_format)
        row += 2

        #Charts
        metal_type_data = metal_data[metal_data['LOC_TYPE'] == sample_type]
        scatter(metal_type_data, sample_type)
        siteLineChart(metal_type_data, sample_type)
        siteScatter(metal_type_data, sample_type)
        
        # add image to sheet
        if sample_type == 'SOIL':
            sheet.insert_image(row, col+1, f'files/{sample_type}.png', {'x_scale': 0.65, 'y_scale': 0.65})
            col += 4
        else:
            sheet.insert_image(row, col+1, f'files/{sample_type}_washed.png', {'x_scale': 0.65, 'y_scale': 0.65})
            sheet.insert_image(row, col+4, f'files/{sample_type}_unwashed.png', {'x_scale': 0.65, 'y_scale': 0.65})
            col += 7
        sheet.insert_image(row, col, f'files/{sample_type}_site_scatter.png', {'x_scale': 0.65, 'y_scale': 0.55})
        sheet.insert_image(row, col+3, f'files/{sample_type}_site_line.png', {'x_scale': 0.65, 'y_scale': 0.65})
        col = 0

        
        row += 15

        #Historic
        historic_stats = stats['historic']
        years = sorted(historic_stats.keys())

        def write_historic(site, ref, ttest, years, subtype, row, col):
            #Site
            sheet.write(row, col, f'Historic - {subtype}', format_bold)
            sheet.write_row(row, col+2, years, format_bold)
            row += 1
            sheet.write(row, col, 'Site Sample Count')
            sheet.write_row(row, col+2, [site[year]['count'] for year in years])
            row += 1
            sheet.write(row, col, 'Site Mean')
            sheet.write_row(row, col+2, [site[year]['mean'] for year in years], number_format)
            row += 1
            sheet.write(row, col, 'Site Standard Deviation')
            sheet.write_row(row, col+2, [site[year]['std'] for year in years], number_format)
            #Now do Ref
            row += 1
            sheet.write(row, col, 'Ref Sample Count')
            sheet.write_row(row, col+2, [ref[year]['count'] for year in years])
            row += 1
            sheet.write(row, col, 'Ref Mean')
            sheet.write_row(row, col+2, [ref[year]['mean'] for year in years], number_format)
            row += 1
            sheet.write(row, col, 'Ref Standard Deviation')
            sheet.write_row(row, col+2, [ref[year]['std'] for year in years], number_format)
            #ttest results
            row += 1
            sheet.write(row, col, 'Site vs Ref T-Test')
            #loop years and get the p test significance
            significance = []
            for year in years:
                if ttest[year]['p']:
                    if ttest[year]['p'] < 0.05:
                        significance.append('Significant')
                    else:
                        significance.append('Not Significant')
                else:
                    significance.append('')
            sheet.write_row(row, col+2, significance)
            row += 1
            sheet.write(row, col, 'Site vs Ref T-Test p-value')
            sheet.write_row(row, col+2, [ttest[year]['p'] for year in years], number_format)
                       
            #group the historic data in excel
            for i in range(8):
                sheet.set_row(row-i, None, None, {'level': 1, 'hidden': True})
            
            row += 1
            return (row, col)

        #Washed
        washed_site = {year: historic_stats[year]['washed_site'] for year in years}
        washed_ref = {year: historic_stats[year]['washed_ref'] for year in years}
        washed_ttest = {year: historic_stats[year]['washed_ttest'] for year in years}
        row, col = write_historic(washed_site, washed_ref, washed_ttest, years, 'Washed', row, col)

        #Unwashed
        row += 1
        col = 0
        unwashed_site = {year: historic_stats[year]['unwashed_site'] for year in years}
        unwashed_ref = {year: historic_stats[year]['unwashed_ref'] for year in years}
        unwashed_ttest = {year: historic_stats[year]['unwashed_ttest'] for year in years}
        row, col = write_historic(unwashed_site, unwashed_ref, unwashed_ttest, years, 'Unwashed', row, col)
        row += 2

        def write_data(metal_df, years, subtype, row, col):
            sheet.write_row(row, col, ['Site Data '+subtype.capitalize() , 'Time Correlation'], format_bold)
            sheet.write_row(row, col+2, years, format_bold)
            row += 1
            #get loc name by combining LOC_ID and LOC_TYPE. Sort by LOC ID

            metal_df['LOC_NAME'] = metal_df['LOC_ID'].astype(str) + ' ' + metal_df['LOC_TYPE'] + ' ' + metal_df['LOC_SUBTYPE']
            metal_df = metal_df.sort_values(by=['LOC_ID'])
            #for each location, write the data

            for loc_name, loc_df in metal_df.groupby('LOC_NAME'):
                #get the time correlation over the years, if we have more than 2 values
                time_corr = loc_df['REPORT_RESULT_VALUE'].corr(loc_df['SAMPLE_YEAR'])
                time_corr = time_corr if not np.isnan(time_corr) else ''
                time_corr = time_corr if (len(loc_df['REPORT_RESULT_VALUE']) > 2) else ''
                #write the location name
                sheet.write(row, col, loc_name)
                sheet.write(row, col+1, time_corr, number_format)
                #write the data by year
                for year in years:
                    mean = loc_df[loc_df['SAMPLE_YEAR'] == year]['REPORT_RESULT_VALUE'].mean()
                    #convert nan to None
                    mean = mean if not np.isnan(mean) else ''
                    sheet.write(row, col+2, mean, number_format)
                    col += 1
                sheet.set_row(row, None, None, {'level': 1, 'hidden': True})

                row += 1
                col = 0
            
            return row


        #Site data washed over time
        metal_type_data = metal_data[metal_data['LOC_TYPE'] == sample_type]
        washed = metal_type_data[metal_type_data['LOC_SUBTYPE'] == 'WASHED']
        row = write_data(washed, years, 'Washed', row, col)

        #site data unwashed over time
        row += 1
        col = 0
        unwashed = metal_type_data[metal_type_data['LOC_SUBTYPE'] == 'UNWASH']
        row = write_data(unwashed, years, 'Unwashed', row, col)
        
        row += 2

    workbook.close()

def siteLineChartSns(metal_data, name='site_line_chart'):
    import seaborn as sns

    #create figure and axis
    fig, ax = plt.subplots()

    #Line chart of site data over time. Each site is a line.
    sns.lineplot(data=metal_data, x='SAMPLE_YEAR', y='REPORT_RESULT_VALUE', hue='LOC_ID', ci=None, ax=ax)

    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.)
    plt.savefig(f'files/{name}.png')
    # plt.show()

    
def siteScatter(metal_data, name='site_scatter'):
    #loc_xname to be the loc_id and first letter of subtype
    metal_data['LOC_XNAME'] = metal_data['LOC_ID'].astype(str) + ' ' + metal_data['LOC_SUBTYPE'].str[0]

    #sort loc_xname
    metal_data = metal_data.sort_values(by=['LOC_XNAME'])
    # show washed in blue, unwashed in orange
    washed = metal_data[metal_data['LOC_SUBTYPE'] == 'WASHED']
    unwashed = metal_data[metal_data['LOC_SUBTYPE'] == 'UNWASH']
    #sort by LOC_XNAME
    washed = washed.sort_values(by=['LOC_XNAME'])
    unwashed = unwashed.sort_values(by=['LOC_XNAME'])
    #X axis should be unique site name. Washed is blue, unwashed is orange.
    sites = metal_data['LOC_XNAME'].unique()

    #sort sites. Note that lox_xname can have a string and number. It should be sorted by number unless there is no number in the string.
    import re
    try:
        sites = sorted(sites, key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
    except:
        sites = sorted(sites)

    fig, ax = plt.subplots()
    plt.xticks(rotation=90)

    for site in sites:
        #on the first site, label the series
        site_unwashed = unwashed[unwashed['LOC_XNAME'] == site]
        site_washed = washed[washed['LOC_XNAME'] == site]
        
        if site == sites[0]:
            ax.scatter(site_unwashed['LOC_XNAME'], site_unwashed['REPORT_RESULT_VALUE'], color='tab:orange', alpha=0.5, label='Unwashed')
            ax.scatter(site_washed['LOC_XNAME'], site_washed['REPORT_RESULT_VALUE'], color='tab:blue', alpha=0.5, label='Washed')
        else:
            ax.scatter(site_unwashed['LOC_XNAME'], site_unwashed['REPORT_RESULT_VALUE'], color='tab:orange', alpha=0.5)
            ax.scatter(site_washed['LOC_XNAME'], site_washed['REPORT_RESULT_VALUE'], color='tab:blue', alpha=0.5)

    # title of chemical name 
    ax.set_title('By Site')
    ax.legend()
    
    plt.savefig(f'files/{name}_site_scatter.png', bbox_inches='tight', dpi=300)
    # plt.show()



def siteLineChart(metal_data, name='site_line_chart'):
    #Line chart of site data over time. Each site is a line.
    sites = metal_data['LOC_ID'].unique()

    #create figure and axis
    fig, ax = plt.subplots()

    for site in sites:
        site_data = metal_data[metal_data['LOC_ID'] == site]
        #we want x axis to be year, but only one data point per year, being the mean.
        years = site_data['SAMPLE_YEAR'].unique()
        #sort years
        years = sorted(years)
        #get the mean of each year
        means = []
        for year in years:
            means.append(site_data[site_data['SAMPLE_YEAR'] == year]['REPORT_RESULT_VALUE'].mean())
        #plot the line
        ax.plot(years, means, label=site)

    #move legend outside of graph area
    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.)
    ax.set_title('Sites Over Time')
    
    plt.savefig(f'files/{name}_site_line.png', bbox_inches='tight', dpi=300)
    # plt.show()
