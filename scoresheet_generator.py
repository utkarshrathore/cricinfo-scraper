#!/usr/bin/env python

# Copyright (c) 2014, Utkarsh Rathore
# All rights reserved.

# Redistribution and use in source and binary forms, with or without 
# modification, are permitted provided that the following conditions are met:

# 1. Redistributions of source code must retain the above copyright notice,
# this list of conditions and the following disclaimer.

# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.

# 3. Neither the name of the copyright holder nor the names of its
# contributors may be used to endorse or promote products derived from
# this software without specific prior written permission.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

__version__ = '1.0'

import os
import sys
import re
import xlwt
import requests
import bs4

"""
Tries to guess the (mostly) unique ESPNcricinfo name of the cricketer as listed 
on the player profile page. Helps me avoid all the crazy player aliases floting around

"""
def get_canonical_player_name(player_str):
	match = re.search(r'view the player profile for (.*)$', player_str)
	if match:
		return match.group(1)
	else:
		return None

""" convert dismissal string into well known dismissal types """
def get_dismissal_type(dismissal_str):
	dismissal_str = dismissal_str.strip()
	if re.search(r'^b',dismissal_str):
		return 'bowled'
	elif re.search(r'^c',dismissal_str):
		return 'caught'
	elif re.search(r'^lbw',dismissal_str):
		return 'lbw'
	elif re.search(r'^run out',dismissal_str):
		return 'runout'
	elif re.search(r'^st',dismissal_str):
		return 'stumped'
	elif re.search(r'^not out',dismissal_str):
		return 'notout'
	else:
		return None

""" Returns a list of batsman stats for the inning """
def parse_bat_inning(bat_inning_soup):
	batsman_stats = []

	try:
		rows = bat_inning_soup.select('.inningsRow')
	except AttributeError:
		return batsman_stats

	for row in rows:
		try:
			player_name = get_canonical_player_name(str(row.select('td > a')[0].get('title')))
			how_out = get_dismissal_type(row.select('.battingDismissal')[0].get_text())
			runs_scored = int(row.select('.battingRuns')[0].get_text())
		except:
			continue

		try:
			strike_rate = float(row.select('td:nth-of-type(9)')[0].get_text())
		except ValueError:
			print '[WARN] Found illegal strike rate for %s -- continuing assuming sr = 0' % player_name
			strike_rate = 0

		batsman_stats.append([player_name, how_out, runs_scored, strike_rate])

	return batsman_stats

""" Returns a list of bowler stats for the inning """
def parse_bowl_inning(bowl_inning_soup):
	bowler_stats = []

	try:
		rows = bowl_inning_soup.select('.inningsRow')
	except AttributeError:
		return bowler_stats

	for row in rows:
		try:
			player_name = get_canonical_player_name(str(row.select('td > a')[0].get('title')))
			overs = float(row.select('td:nth-of-type(3)')[0].get_text())
			wickets = float(row.select('td:nth-of-type(6)')[0].get_text())
			economy = float(row.select('td:nth-of-type(7)')[0].get_text())
		except:
			print '[WARN] Unhandled exception in bowling stats for %s. Please try adding stats manually' % player_name
			continue

		bowler_stats.append([player_name, overs, wickets, economy])

	return bowler_stats

""" Write batting stats for the inning to excel worksheet """
def write_bat_stats_to_excel(batsman_stats, team, batting_sheet, row):
	for batsman_stat in batsman_stats:
		batting_sheet.write(row, 0, batsman_stat[0])
		batting_sheet.write(row, 1, team)
		batting_sheet.write(row, 2, batsman_stat[1])
		batting_sheet.write(row, 4, batsman_stat[2])
		batting_sheet.write(row, 5, batsman_stat[3])
		if batsman_stat[1] in ['bowled', 'notout', 'lbw']:
			row += 1
			continue
		else:
			batting_sheet.write(row, 3, 'TBD')
		row += 1

	return row

""" Write bowling stats for the inning to excel worksheet """
def write_bowl_stats_to_excel(bowler_stats, team, bowling_sheet, row):
	for bowler_stat in bowler_stats:
		bowling_sheet.write(row, 0, bowler_stat[0])
		bowling_sheet.write(row, 1, team)
		bowling_sheet.write(row, 2, bowler_stat[1])
		bowling_sheet.write(row, 3, bowler_stat[2])
		bowling_sheet.write(row, 4, bowler_stat[3])
		row += 1

	return row

if __name__ == "__main__":
	filename = os.path.basename(__file__).split('.')[0]

	try:
		url = sys.argv[1]
	except:
		print 'USAGE: python %s <Scorecard URL>' % filename
		sys.exit(2)

	try:
		response = requests.get(url)
	except:
		print '[FATAL] URL fetch failed. Please check the URL (Sample: http://www.espncricinfo.com/...)'
		sys.exit(2)

	soup = bs4.BeautifulSoup(response.text)

	bat_inning_1_soup = soup.find('table', class_='inningsTable', attrs={'id':'inningsBat1'})
	bat_inning_2_soup = soup.find('table', class_='inningsTable', attrs={'id':'inningsBat2'})
	bowl_inning_1_soup = soup.find('table', class_='inningsTable', attrs={'id':'inningsBowl1'})
	bowl_inning_2_soup = soup.find('table', class_='inningsTable', attrs={'id':'inningsBowl2'})

	team1 = str(bat_inning_1_soup.find('tr', class_="inningsHead").find('td', attrs={"colspan":"2"}).get_text().split()[0])
	team2 = str(bat_inning_2_soup.find('tr', class_="inningsHead").find('td', attrs={"colspan":"2"}).get_text().split()[0])

	'''
	Teams can alternatively be queries as below but above is more robust
	team1 = str(bat_inning_1.find('tr', class_="inningsHead").select('td')[1].get_text().split()[0])
	team2 = str(bat_inning_2.find('tr', class_="inningsHead").select('td')[1].get_text().split()[0])
	'''

	bat_inning_1_stats = parse_bat_inning(bat_inning_1_soup)
	bat_inning_2_stats = parse_bat_inning(bat_inning_2_soup)
	bowl_inning_1_stats = parse_bowl_inning(bowl_inning_1_soup)
	bowl_inning_2_stats = parse_bowl_inning(bowl_inning_2_soup)

	workbook = xlwt.Workbook()
	
	batting_sheet = workbook.add_sheet('Batting')
	batting_sheet.write(0, 0, 'player_name')
	batting_sheet.write(0, 1, 'team')
	batting_sheet.write(0, 2, 'how_out')
	batting_sheet.write(0, 3, 'fielder_involved')
	batting_sheet.write(0, 4, 'runs_scored')
	batting_sheet.write(0, 5, 'batting_strike_rate')
	batting_sheet_row = 1

	if bat_inning_1_stats:
		batting_sheet_row = write_bat_stats_to_excel(bat_inning_1_stats, team1, batting_sheet, batting_sheet_row)
	else:
		print '[WARN] No data for Batting Inning 1 -- skipping'

	if bat_inning_2_stats:
		batting_sheet_row = write_bat_stats_to_excel(bat_inning_2_stats, team2, batting_sheet, batting_sheet_row)
	else:
		print '[WARN] No data for Batting Inning 2 -- skipping'

	bowling_sheet = workbook.add_sheet('Bowling')
	bowling_sheet.write(0, 0, 'player_name')
	bowling_sheet.write(0, 1, 'team')
	bowling_sheet.write(0, 2, 'overs')
	bowling_sheet.write(0, 3, 'wickets')
	bowling_sheet.write(0, 4, 'economy')
	bowling_sheet_row = 1

	if bowl_inning_1_stats:
		bowling_sheet_row = write_bowl_stats_to_excel(bowl_inning_1_stats, team2, bowling_sheet, bowling_sheet_row)
	else:
		print '[WARN] No data for Bowling Inning 1 -- skipping'

	if bowl_inning_2_stats:
		bowling_sheet_row = write_bowl_stats_to_excel(bowl_inning_2_stats, team1, bowling_sheet, bowling_sheet_row)
	else:
		print '[WARN] No data for Bowling Inning 2 -- skipping'

	# All done. Save the workbook.
	workbook_name = team1 + ' ' + 'v' + ' ' + team2 + '.xls'
	workbook.save(workbook_name)
	print '[ALL DONE] Scoresheet saved as %s' % workbook_name
