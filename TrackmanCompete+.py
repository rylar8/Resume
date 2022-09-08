import sqlite3

#Don't have: Runners in scoring position or tying run at the plate or on base in 7th or later


def openDatabase():
    path = "c:\\Users\\rlars\Desktop\Knights Baseball\Project Compete+\Project Compete+ Database"
    con = sqlite3.connect(path)
    cur = con.cursor()

    return cur, con

#Hitter leverage strike
def accumHitterLeverage():
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult FROM ProjectCompetePlus WHERE balls > strikes')
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    for tup in allPitches:
        if tup[0] in strikeList:
            success += 1
        total += 1
    
    return success / total

#first pitch strike to batters
def accumFPS():
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult FROM ProjectCompetePlus WHERE pitchPA = 1')
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    for tup in allPitches:
        if tup[0] in strikeList:
            success += 1
        total += 1
    
    return success / total

#strikes to first batter you face
def accumFirstBatter():
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, pitcher, hitter, date FROM ProjectCompetePlus')
    allPitches = cur.fetchall()

    dates = {}
    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    for pitch in allPitches:
        pitcher = pitch[1]
        hitter = pitch[2]
        date = pitch[3]
        if date in dates:
            if pitcher in dates[date]:
                if hitter == dates[date][pitcher]:
                    if pitch[0] in strikeList:
                        success += 1
                    total += 1 
            else:
                dates[date][pitcher] = hitter
        else:
            dates[date] = {pitcher : hitter}

    return success / total

#strike after negative result
def accumNegativeResult():
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, outcome, KorBB, inning, paInning, pitcher FROM ProjectCompetePlus')
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
    return success / total

#strike after 20+ pitches in inning
def accum20Pitches():
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, pitcher, date, inning FROM ProjectCompetePlus')
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    used = []
    for pitch in allPitches:
        pitcher = pitch[1]
        date = pitch[2]
        inning = pitch[3]

        if (pitcher, date, inning) in used:
            continue
        cur.execute(f'SELECT pitchResult FROM ProjectCompetePlus WHERE pitcher = ? AND date = ? AND inning = ?', (pitcher, date, inning))
        pitches = cur.fetchall()
        used.append((pitcher, date, inning))

        if len(pitches) > 20:
            pitchesAfter20 = pitches[20:]
            for tup in pitchesAfter20:
                if tup[0] in strikeList:
                    success += 1
                total += 1

    return success / total

#strike after 75+ pitches in outing
def accum75Pitches():
    cur, con = openDatabase()

    cur.execute(f'SELECT pitchResult, pitcher, date FROM ProjectCompetePlus')
    allPitches = cur.fetchall()

    total = 0
    success = 0
    strikeList = ['FoulBall', 'InPlay', 'StrikeCalled', 'StrikeSwinging']
    used = []
    for pitch in allPitches:
        pitcher = pitch[1]
        date = pitch[2]

        if (pitcher, date) in used:
            continue
        cur.execute(f'SELECT pitchResult FROM ProjectCompetePlus WHERE pitcher = ? AND date = ?', (pitcher, date))
        pitches = cur.fetchall()
        used.append((pitcher, date))

        if len(pitches) > 75:
            pitchesAfter75 = pitches[75:]
            for tup in pitchesAfter75:
                if tup[0] in strikeList:
                    success += 1
                total += 1

    return success / total

#WCL Averages
#Hitter Leverage : 0.6328491548300019
#FPS : 0.5689623440203188
#First Batter Faced: 0.6038945618617906
#Good After Bad : 0.599106431614311
#After 20+ : 0.6093080220453154
#After 75+ : 0.62468999724442

