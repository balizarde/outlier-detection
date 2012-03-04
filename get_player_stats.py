# ----------------------------------------------------
# Collect player statistics from http://www.databasebasketball.com/
# 
# Sercan Taha Ahi, Nov 2011 (tahaahi at gmail dot com)
# ----------------------------------------------------
import urllib, urllib2
import string
from pyExcelerator.Workbook import *
from pyExcelerator.Style import *

def parse_comp(key, year, n_active_players, ws0):
    url_list = 'http://www.databasebasketball.com/players/playerlist.htm?lt=' + key
    try:
        page = urllib2.urlopen(url_list).read()
    except:
        print 'Error: ' + key
        return 0
        
    ctr = page.count('<td><a href="playerpage.htm?ilkid=')
    print 'Surnames starting with ' + key + ': \t' + str(ctr)
    
    ptr1 = 0
    str1 = '<td><a href="playerpage.htm?ilkid='
    str2 = '">'
    str3 = '</a>'
    str4 = '<td class=cen>'
    str5 = '</td>'
    
    strp1 = 'Regular Season Stats'
    strp2 = '<td align="center">'
    strp3 = '</td>'
    strp4 = '<td align="right">'
    for i in range(0,ctr):
        ptr1 = page.find(str1, ptr1)
        ptr2 = page.find(str2, ptr1)
        ptr3 = page.find(str3, ptr2)
        ptr4 = page.find(str4, ptr3)
        ptr5 = page.find(str5, ptr4)
        
        player_code = page[ptr1+len(str1):ptr2]
        
        player_name = page[ptr2+len(str2):ptr3];
        player_name = player_name.replace('<b>', '')
        player_name = player_name.replace('</b>', '')
        player_name = player_name.replace('&nbsp;', ' ')
        
        player_years = page[ptr4+len(str4):ptr5]
        player_year1 = int(player_years[0:4])
        player_year2 = int(player_years[7:12])
        
        if player_year1<= year and year+1 <= player_year2:
            ptr6 = page.find(str4, ptr5)
            ptr7 = page.find(str5, ptr6)
            player_position = page[ptr6+len(str4):ptr7]
            
            url_player = 'http://www.databasebasketball.com/players/playerpage.htm?ilkid=' + player_code
            try:
                page_player = urllib2.urlopen(url_player).read()
            except:
                print 'Error: Cannot access player page, ' + player_name
                break
            
            tmp = str(year+1)
            target_season = str(year) + '-' + tmp[2:4]
            ptrp1 = 0
            
            # Move to regular season stats
            ptrp1 = page_player.find(strp1, ptrp1)
            
            # Move to target season
            ptrp1 = page_player.find(target_season, ptrp1)
            
            if ptrp1 > 0:
                n_active_players = n_active_players + 1
                
                # Age
                ptrp1 = page_player.find(strp2, ptrp1)
                ptrp2 = page_player.find(strp3, ptrp1)
                player_age = int(page_player[ptrp1+len(strp2):ptrp2])
                
                # Team name
                tstr = 'yr=' + str(year) + '">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</a>', ptrp1)
                player_team = page_player[ptrp1+len(tstr):ptrp2]
                
                # Total played games
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_n_games = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Total played minutes
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_min = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Minutes per game
                player_minpg = float(player_min) / float(player_n_games)
                
                # Points
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_pts = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Points per game
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_ptspg = float(page_player[ptrp1+len(tstr):ptrp2])
                
                # Field goal made
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_fgm = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Field goal attempted
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_fga = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Field goal percentage
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_fgp = float(page_player[ptrp1+len(tstr):ptrp2])
                
                # Free throw made
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_ftm = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Free throw attemted
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_fta = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Free throw percentage
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_ftp = float(page_player[ptrp1+len(tstr):ptrp2])
                
                # 3-point made
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_3pm = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # 3-point attempted
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_3pa = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # 3-point percentage
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_3pp = float(page_player[ptrp1+len(tstr):ptrp2])
                
                # Offensive rebounds
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_orb = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Offensive rebounds per game
                player_orbpg = float(player_orb) / float(player_n_games)
                
                # Defensive rebounds
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_drb = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Offensive rebounds per game
                player_drbpg = float(player_drb) / float(player_n_games)
                
                # Total Rebounds
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_trb = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Total rebounds per game
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_trbpg = float(page_player[ptrp1+len(tstr):ptrp2])
                
                # Number of assists
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_ast = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Assists per game
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_astpg = float(page_player[ptrp1+len(tstr):ptrp2])
                
                # Number of steals
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_stl = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Steals per game
                player_stlpg = float(player_stl) / float(player_n_games)
                
                # Number of blocks
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_blk = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Blocks per game
                player_blkpg = float(player_blk) / float(player_n_games)
                
                # Number of turnovers
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_to = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Turnovers per game
                player_topg = float(player_to) / float(player_n_games)
                
                # Personal fouls
                tstr = '<td align="right">'
                ptrp1 = page_player.find(tstr, ptrp2)
                ptrp2 = page_player.find('</td>', ptrp1)
                player_pf = int(page_player[ptrp1+len(tstr):ptrp2])
                
                # Personal fouls per game
                player_pfpg = float(player_pf) / float(player_n_games)
                
                # Print the data
                ws0.write(n_active_players, 0, player_name)
                ws0.write(n_active_players, 1, player_code)
                ws0.write(n_active_players, 2, player_position)
                ws0.write(n_active_players, 3, player_team)
                ws0.write(n_active_players, 4, player_year1)
                ws0.write(n_active_players, 5, player_year2)
                ws0.write(n_active_players, 6, player_age)
                ws0.write(n_active_players, 7, player_n_games)
                ws0.write(n_active_players, 8, player_min)
                ws0.write(n_active_players, 9, player_minpg)
                ws0.write(n_active_players, 10, player_pts)
                ws0.write(n_active_players, 11, player_ptspg)
                ws0.write(n_active_players, 12, player_fgm)
                ws0.write(n_active_players, 13, player_fga)
                ws0.write(n_active_players, 14, player_fgp)
                ws0.write(n_active_players, 15, player_ftm)
                ws0.write(n_active_players, 16, player_fta)
                ws0.write(n_active_players, 17, player_ftp)
                ws0.write(n_active_players, 18, player_3pm)
                ws0.write(n_active_players, 19, player_3pa)
                ws0.write(n_active_players, 20, player_3pp)
                ws0.write(n_active_players, 21, player_orb)
                ws0.write(n_active_players, 22, player_orbpg)
                ws0.write(n_active_players, 23, player_drb)
                ws0.write(n_active_players, 24, player_drbpg)
                ws0.write(n_active_players, 25, player_trb)
                ws0.write(n_active_players, 26, player_trbpg)
                ws0.write(n_active_players, 27, player_ast)
                ws0.write(n_active_players, 28, player_astpg)
                ws0.write(n_active_players, 29, player_stl)
                ws0.write(n_active_players, 30, player_stlpg)
                ws0.write(n_active_players, 31, player_blk)
                ws0.write(n_active_players, 32, player_blkpg)
                ws0.write(n_active_players, 33, player_to)
                ws0.write(n_active_players, 34, player_topg)
                ws0.write(n_active_players, 35, player_pf)
                ws0.write(n_active_players, 36, player_pfpg)
                
                ostr = str(n_active_players).rjust(4) + '\t' + player_name
                print ostr
        
        ptr1 = ptr5
        
    return n_active_players

    
def write_categories(ws0):
    ws0.write(0, 0, 'Player')
    ws0.write(0, 1, 'Code')
    ws0.write(0, 2, 'Position')
    ws0.write(0, 3, 'Team')
    ws0.write(0, 4, 'From')
    ws0.write(0, 5, 'Until')
    ws0.write(0, 6, 'Age')
    ws0.write(0, 7, 'G')
    ws0.write(0, 8, 'MIN')
    ws0.write(0, 9, 'MINPG')
    ws0.write(0, 10, 'PTS')
    ws0.write(0, 11, 'PTSPG')
    ws0.write(0, 12, 'FGM')
    ws0.write(0, 13, 'FGA')
    ws0.write(0, 14, 'FGP')
    ws0.write(0, 15, 'FTM')
    ws0.write(0, 16, 'FTA')
    ws0.write(0, 17, 'FTP')
    ws0.write(0, 18, '3PM')
    ws0.write(0, 19, '3PA')
    ws0.write(0, 20, '3PP')
    ws0.write(0, 21, 'ORB')
    ws0.write(0, 22, 'ORBPG')
    ws0.write(0, 23, 'DRB')
    ws0.write(0, 24, 'DRBPG')
    ws0.write(0, 25, 'TRB')
    ws0.write(0, 26, 'TRBPG')
    ws0.write(0, 27, 'AST')
    ws0.write(0, 28, 'ASTPG')
    ws0.write(0, 29, 'STL')
    ws0.write(0, 30, 'STLPG')
    ws0.write(0, 31, 'BLK')
    ws0.write(0, 32, 'BLKPG')
    ws0.write(0, 33, 'TO')
    ws0.write(0, 34, 'TOPG')
    ws0.write(0, 35, 'PF')
    ws0.write(0, 36, 'PFPG')

def main():
    if len(sys.argv) == 1:
        target_year = 1997
    else:
        target_year = int(sys.argv[1])
    
    print "Collecting statistics for the " + str(target_year) + '-' + str(target_year+1) + " season...\n"
    fname = str(target_year) + '-' + str(target_year+1) + '.xls'
    style = XFStyle()
    wb = Workbook()
    ws0 = wb.add_sheet('stats')
    write_categories(ws0)
    
    n_players = 0
    for key in range(ord('a'), ord('z')+1):
        n_players = parse_comp(chr(key), target_year, n_players, ws0)
    
    print '# of players = ' + str(n_players)
    wb.save(fname)

    
main()
