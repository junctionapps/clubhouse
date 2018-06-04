#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Author: aaron@junctionapps.ca
# Project: Clubhouse utilities sample
# Clubhouse is a project management platform for software teams that provides
# the perfect balance of simplicity and structure. https://clubhouse.io/

# Copyright 2018 Junction Applications Limited
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import os
import html
import requests
import xlsxwriter


CH_TOKEN = os.getenv('CLUBHOUSE_TOKEN', None)
CH_BASE = 'https://api.clubhouse.io/'
CH_API_VER = 'api/v2/'


def search_dict(entities, search_in, search_val, return_key):
    """
    Utility: Rudimentary search through a list of dictionaries. First find wins
    :param entities: a list of entities from Clubhouse (JSON)
    :param search_in: the name of the dictionary key to search
    :param search_val: the value to look for in the search_in key
    :param return_key: the dictionary key to return
    :return: the value in the dictionary's return_key key
    """
    # https://stackoverflow.com/a/7079297
    return_item = next((item for item in entities if item[search_in] == search_val), None)
    if return_item:
        return return_item.get(return_key, None)
    else:
        return None


def url(entity):
    """
    Returns a properly formatted url given module's constants
    entity is one of the Clubhouse entities (epics, labels, stories etc.)
    :param entity: a string with the entity type
    :return: url as a string for that entity
    """
    return '{ch_base}{ch_ver}{ch_entity}?token={ch_token}'.format(ch_base=CH_BASE,
                                                                  ch_ver=CH_API_VER,
                                                                  ch_entity=entity,
                                                                  ch_token=CH_TOKEN)


def entity_list(entity):
    """
    :param entity: a string with the entity type
    :return: a JSON packet listing of entities
    """
    response = requests.get(url(entity))
    return response.json()


def search_stories(query, page_size=25):
    """
    Returns the first 1000 results from search stories.
    :param query: query in the form of a proper Clubhouse query string outlined at:
                  https://help.clubhouse.io/hc/en-us/articles/360000046646-Search-Operators
                  note that spaces in labels should be replaced by hyphens
    :param page_size: the number of search results to return in a single call. Clubhouse has
                      has a default limit of 25 for performance purposes
    :return: dictionary with keys data:  holds the list of stories,
                                  total: the total number of search results reported by Clubhouse
                                  warning (optional): exists if more than 1000 records were found
    """
    #
    #
    body_params = {'page_size': page_size, 'query': query}
    entity = 'search/stories'

    # get the first results
    data = dict()
    first_page = requests.get(url(entity), body_params).json()
    stories = first_page.get('data', None)
    next_page_token = first_page.get('next', None)
    total_stories = first_page.get('total', None)
    while_count = 0

    while next_page_token:
        while_count += 1
        # if we're calling more than this we've probably made a mistake or need to make an adjustment to our
        # query string
        # picking arbitrary 1000 limit hopefully keeping us under the Clubhouse 200/min rate
        # may be better to simply look at total prior to entering while loop
        if while_count * page_size > 1000:
            data.update({'warning': 'Excessive api calls resulted in truncated data set. '
                                    'About {rc} of {tc} returned'.format(rc=page_size*while_count, tc=total_stories)})
            break
        # so we strip off that last / because the next page token 'next' url starts with a /
        np_url = '{base}{npt}&token={apitoken}'.format(base=CH_BASE[:-1],
                                                       npt=html.unescape(next_page_token),
                                                       apitoken=CH_TOKEN)
        next_page = requests.get(np_url).json()
        stories.extend(next_page['data'])
        next_page_token = next_page.get('next', None)

    data.update({'data': stories, 'total': total_stories})
    return data


def epics_with_label(epics, search_label):
    matching_epics = dict()
    for epic in epics:
        for label in epic['labels']:
            if label['name'] == search_label:
                matching_epics.update({epic['id']: epic['name']})
    return matching_epics


def output_stories(stories, output_filename):
    """
    :param stories: list of stories
    :param output_filename: the file to save the results to. Should be format name.xlsx
    :return: Nothing
    """

    # we need to look up the epics as we need an id
    # for this demonstration, we'll use the epics as the sheet names in excel
    all_epics = entity_list('epics')
    epics = epics_with_label(epics=all_epics, search_label='Test Suite')

    # we need an index for the next row to write to in each sheet
    next_row_index = {}

    # create the workbook
    workbook = xlsxwriter.Workbook(output_filename)

    headings = ['ID',
                'Process',
                'Test Description',
                'Pass/Fail',
                'Comments - Include URL when reporting an on-screen issue', ]
    text_wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    first_row_format = workbook.add_format({'bold': True})
    df = workbook.add_format({'valign': 'top'})  # default_format

    for epic_id, epic_name in epics.items():
        # we need to now the next row for each type
        next_row_index[epic_id] = 0
        worksheet = workbook.add_worksheet(name=epic_name)
        worksheet.set_column(1, 1, 20)  # Set column widths
        worksheet.set_column(2, 2, 70)
        worksheet.set_column(3, 3, 40)
        for column, heading in enumerate(headings):
            worksheet.write(0, column, heading, first_row_format)

    for story in stories:
        ws_name = epics[story['epic_id']]    # the sheet name is the value in epics dictionary
        ws = workbook.get_worksheet_by_name(ws_name)
        next_row_index[story['epic_id']] += 1
        ws.write(next_row_index[story['epic_id']], 0, story['id'], df)
        ws.write(next_row_index[story['epic_id']], 1, story['name'], df)
        ws.write(next_row_index[story['epic_id']], 2, '\n'.join(story['description'].splitlines()), text_wrap_format)

    try:
        workbook.close()
    except PermissionError:
        print('Check that the output file is not opened in Excel or other application.\n'
              'We could not write to it due to a permission denied error.')


def copy_stories_to_excel(query, output_to):
    """

    :param query: query in the form of a proper Clubhouse query string outlined at:
                  https://help.clubhouse.io/hc/en-us/articles/360000046646-Search-Operators
    :param output_to: the file to save the results to. Should be format name.xlsx
    :return: Nothing
    """
    """ Queries the stories, outputs to console and creates the excel file"""
    stories = search_stories(query=query)
    print('Total stories with label:To-be-tested: {ts}'.format(ts=stories['total']))
    print('Sending {} stories to output file'.format(len(stories['data'])))
    output_stories(stories['data'], output_to)


def main():
    # we'll search for any stories with the label "To be tested"
    # we could get more complicated with combining search terms:
    # example: get all stories in the "Test scripts" project that are not complete and have no owner
    # query = "project:Test-scripts !state:complete !has:owner "
    # for this sample, we'll just get the ones across all projects with label: To be tested
    query = 'label:To-be-tested'
    copy_stories_to_excel(query=query, output_to='output_sample.xlsx')


if __name__ == '__main__':
    if CH_TOKEN:
        main()
    else:
        print('The CLUBHOUSE_TOKEN environment variable is not set.\n'
              'See https://help.clubhouse.io/hc/en-us/articles/205701199-Clubhouse-API-Tokens\n'
              'Then make the token available as an environment variable (usually in your virtual environment\n')
