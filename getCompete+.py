import sqlite3

def openDatabase():
    path = "c:\\Users\\rlars\Desktop\Knights Baseball\Postgame Reports\Databases\Knights2022"
    con = sqlite3.connect(path)
    cur = con.cursor()

    return cur, con

#Don't have: Runners in scoring position or tying run at the plate or on base in 7th or later

#Compete+ = Actual Strikes Thrown / League Expected Strikes Thrown

#strikes thrown in hitter leverage count
def getHitterLeverage(pitcher, tableName):
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult FROM {tableName} WHERE balls > strikes AND pitcher = ?', (pitcher,))
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    for tup in allPitches:
        if tup[0] in strikeList:
            success += 1
        total += 1
    
    return success , total

#first pitch strike
def getFPS(pitcher, tableName):
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult FROM {tableName} WHERE pitchPA = 1 AND pitcher = ?', (pitcher,))
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    for tup in allPitches:
        if tup[0] in strikeList:
            success += 1
        total += 1
    
    return success , total

#strikes to first batter faced
def getFirstBatter(pitcher, tableName):
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, hitter, inning FROM {tableName} WHERE pitcher = ?', (pitcher,))
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    hitter = allPitches[0][1]
    i = 0
    while allPitches[i][1] == hitter:
        if allPitches[i][0] in strikeList:
            success += 1
        total += 1
        i += 1

    return success , total

#strike after negative result
def getNegativeResult(pitcher, tableName):
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, outcome, KorBB, inning, paInning, pitcher FROM {tableName} WHERE pitcher = ?', (pitcher,))
    allPitches = cur.fetchall()
    

    inning = 0
    paInning = 0
    pitcher = ''
    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    hitList = ['Single', 'Double', 'Triple', 'HomeRun']

    for pitch in allPitches:
        if pitch[1] in hitList or pitch[2] == 'Walk':
            inning, paInning, pitcher = pitch[3], int(pitch[4]) + 1 , pitch[5]
            continue
        if (pitch[3], pitch[4], pitch[5]) == (inning, paInning, pitcher):
            if pitch[0] in strikeList:
                success += 1
            total += 1
    return success , total

#strike after 20+ pitches in outing
def get20Pitches(pitcher, tableName):
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, pitcher, inning FROM {tableName} WHERE pitcher = ?', (pitcher,))
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    used = []
    for pitch in allPitches:
        inning = pitch[2]

        if (pitcher, inning) in used:
            continue
        cur.execute(f'SELECT pitchResult FROM {tableName} WHERE pitcher = ? AND inning = ?', (pitcher, inning))
        pitches = cur.fetchall()
        used.append((pitcher, inning))

        if len(pitches) > 20:
            pitchesAfter20 = pitches[20:]
            for tup in pitchesAfter20:
                if tup[0] in strikeList:
                    success += 1
                total += 1

    return success , total

#strike after 75+ pitches in outing
def get75Pitches(pitcher, tableName):
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult FROM {tableName} WHERE pitcher = ?', (pitcher,))
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    if len(allPitches) > 75:
        pitchesAfter75 = allPitches[75:]
        for tup in pitchesAfter75:
            if tup[0] in strikeList:
                success += 1
            total += 1

    return success , total

def getCompetePlus(pitcher, tableName):

    leagueStandard = {'hitterLeverage' : 0.6328491548300019, 'FPS' : 0.5689623440203188, 'firstBatterFaced': 0.6038945618617906,
    'goodAfterBad' : 0.599106431614311, 'after20+' : 0.6093080220453154, 'after75+' : 0.62468999724442}

    leagueTotalSuccess = (leagueStandard['hitterLeverage'] * getHitterLeverage(pitcher, tableName)[1] + leagueStandard['FPS'] * getFPS(pitcher, tableName)[1] + 
    leagueStandard['firstBatterFaced'] * getFirstBatter(pitcher, tableName)[1] + leagueStandard['goodAfterBad'] * getNegativeResult(pitcher, tableName)[1]
    + leagueStandard['after20+'] * get20Pitches(pitcher, tableName)[1] + leagueStandard['after75+'] * get75Pitches(pitcher, tableName)[1])

    pitcherTotalSuccess = (getHitterLeverage(pitcher, tableName)[0]  + getFPS(pitcher, tableName)[0] + getFirstBatter(pitcher, tableName)[0] 
    + getNegativeResult(pitcher, tableName)[0] + get20Pitches(pitcher, tableName)[0] + get75Pitches(pitcher, tableName)[0])

    pitcherTotal = (getHitterLeverage(pitcher, tableName)[1]  + getFPS(pitcher, tableName)[1] + getFirstBatter(pitcher, tableName)[1] 
    + getNegativeResult(pitcher, tableName)[1] + get20Pitches(pitcher, tableName)[1] + get75Pitches(pitcher, tableName)[1])


    return int(round(pitcherTotalSuccess / leagueTotalSuccess, 2) * 100), pitcherTotalSuccess, leagueTotalSuccess

def iterateOverGames(pitchers):

    games = [('2022-06-03', 'YakimaValley'), ('2022-06-04', 'YakimaValley'),('2022-06-05', 'YakimaValley'),
    ('2022-06-14', 'Cowlitz'),('2022-06-15', 'Cowlitz'),('2022-06-16', 'Cowlitz'),('2022-06-17', 'YakimaValley'),('2022-06-18', 'YakimaValley'),('2022-06-19', 'YakimaValley'),
    ('2022-06-21', 'WallaWalla'),('2022-06-22', 'WallaWalla'),('2022-06-23', 'WallaWalla'),('2022-06-24', 'Bellingham'),('2022-06-25', 'Bellingham'),('2022-06-26', 'Bellingham'),
    ('2022-06-28', 'Springfield'),('2022-06-29', 'Springfield'),('2022-06-30', 'Springfield'),('2022-07-01', 'PortAngeles'),('2022-07-02', 'PortAngeles'),('2022-07-03', 'PortAngeles'),
    ('2022-07-04', 'Portland'),('2022-07-05', 'Ridgefield'),('2022-07-06', 'Ridgefield'),('2022-07-07', 'Ridgefield'),('2022-07-08', 'Bend'),('2022-07-09', 'Bend'),('2022-07-10', 'Bend'),
    ('2022-07-13', 'Edmonton'),('2022-07-15', 'Wenatchee'),('2022-07-16', 'Wenatchee'),('2022-07-17', 'Wenatchee'),
    ('2022-07-18', 'Portland'),('2022-07-19', 'Cowlitz'),('2022-07-20', 'Cowlitz'),('2022-07-21', 'Cowlitz'),('2022-07-22', 'Portland'),('2022-07-23', 'Portland'),('2022-07-24', 'Portland'),
    ('2022-07-25', 'Portland'),('2022-07-26', 'Springfield'),('2022-07-27', 'Springfield'),('2022-07-28', 'Springfield'),('2022-07-29', 'WallaWalla'),('2022-07-30', 'WallaWalla'),('2022-07-31', 'WallaWalla'),
    ('2022-08-02', 'Bend'),('2022-08-03', 'Bend'),('2022-08-04', 'Bend'),('2022-08-05', 'Ridgefield'),('2022-08-06', 'Ridgefield'),('2022-08-07', 'Ridgefield')]

    months = {'01' : 'January', '02': 'February' , '03': 'March' , '04': 'April' , '05' : 'May' , '06': 'June', '07': 'July', '08': 'August', '09': 'September', '10': 'October', '11': 'November', '12':'December'}

    competeDic = {}
    for game in games:
        date = game[0].split('-')
        tableName = f'{months[date[1]]}{date[2]}v{game[1]}'

        for pitcher in pitchers:
            try:
                competePlus, pitcherSuccess, leagueSuccess = getCompetePlus(pitcher, tableName)
                if pitcher in competeDic:
                    competeDic[pitcher][0] += pitcherSuccess
                    competeDic[pitcher][1] += leagueSuccess
                else:
                    competeDic[pitcher] = (pitcherSuccess, leagueSuccess)
            except:
                pass
    
    print('Corvallis Knights Compete+')
    for pitcher in competeDic:
        print(f'{pitcher}: {int(round(competeDic[pitcher][0] / competeDic[pitcher][1], 2) * 100)}')

pitchers = ['Brotherton, Duke', 'Clark, Will', 'Day, Cameron', 'Deschryver, Nathan', 'Feist, Neil', 'Gartrell, Joey', 'Haider, Rylan', 'Kantola, Kaleb', 
'Lawson, Ian', 'Marshall, Nathan', 'Maylett, Brady', 'Quinn, Victor', 'Ross, Ethan', 'Scott, Matt', 'Segel, Kaden', 'Wiese, Sean']

iterateOverGames(pitchers)
