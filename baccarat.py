# This is a python file to show how the game works
import random
import csv

CARDS = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'] * 32
OUTCOME = ['Player wins', 'Banker wins', 'Tie']


# Inclusive range function
irange = lambda start, end: range(start, end + 1)


def compute_score(hand):
    """Compute the score of a hand"""
    total_value = 0
    for card in hand:
        if card in ['J', 'Q', 'K', '10']:
            total_value += 0
        elif card == 'A':
            total_value += 1
        else:
            total_value += int(card)
        
    return total_value % 10

def play():
    """Returns the winner"""
    player_hand = []
    banker_hand = []

    CARDS.remove(random.choice(CARDS))

    # get new card remove the card from the deck
    player_hand.append(random.choice(CARDS))
    CARDS.remove(player_hand[0])

    banker_hand.append(random.choice(CARDS))
    CARDS.remove(banker_hand[0])

    player_hand.append(random.choice(CARDS))
    CARDS.remove(player_hand[1])

    banker_hand.append(random.choice(CARDS))
    CARDS.remove(banker_hand[1])

    player_score = compute_score(player_hand)
    banker_score = compute_score(banker_hand)

    print('閒家:\t' + player_hand[0] + '\t' + player_hand[1])
    print('閒家點數\t' + str(player_score))
    print('莊家:\t' + banker_hand[0] + '\t' + banker_hand[1])
    print('莊家點數\t' + str(banker_score))

    # Natural
    if player_score in [8, 9] or banker_score in [8, 9]:
        print('-' * 20)
        print('閒家最後點數\t' + str(player_score))
        print('莊家最後點數\t' + str(banker_score))
        print('*' * 20)

        if player_score != banker_score:
            return OUTCOME[banker_score > player_score]
        else:
            return OUTCOME[2]

    # Player has low score
    if player_score in irange(0, 5):
        # Player get's a third card
        player_hand.append(random.choice(CARDS))
        CARDS.remove(player_hand[2])
        player_third = compute_score([player_hand[2]])
        print('閒家博牌:\t' + player_hand[2])

        # Determine if banker needs a third card
        if (banker_score == 6 and player_third in [6, 7]) or \
           (banker_score == 5 and player_third in irange(4, 7)) or \
           (banker_score == 4 and player_third in irange(2, 7)) or \
           (banker_score == 3 and player_third != 8) or \
           (banker_score in [0, 1, 2]):
            banker_hand.append(random.choice(CARDS))
            CARDS.remove(banker_hand[2])
            print('莊家博牌:\t' + banker_hand[2])

    elif player_score in [6, 7]:
        if banker_score in irange(0, 5):
            banker_hand.append(random.choice(CARDS))
            CARDS.remove(banker_hand[2])
            print('莊家博牌:\t' + banker_hand[2])

    # Compute the scores again and return the outcome
    player_score = compute_score(player_hand)
    banker_score = compute_score(banker_hand)

    print('-' * 20)
    print('閒家最後點數\t' + str(player_score))
    print('莊家最後點數\t' + str(banker_score))
    print('*' * 20)

    if player_score != banker_score:
        return OUTCOME[banker_score > player_score]
    else:
        return OUTCOME[2]

# print(play())



for t in range(20):
    results = []
    
    for _ in range(100):
        ## generate new deck
        CARDS = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'] * 32

        player_wins = 0
        banker_wins = 0
        ties = 0
        tie_index = []
        
        for i in range(60):
            outcome = play()
            if outcome == 'Player wins':
                player_wins += 1
            elif outcome == 'Banker wins':
                banker_wins += 1
            else:
                ties += 1
                tie_index.append(i)



        print('閒家贏:\t' + str(player_wins))
        print('莊家贏:\t' + str(banker_wins))

        differences = [tie_index[i+1] - tie_index[i] for i in range(len(tie_index)-1)]
        print('和局:\t' + str(ties) + '\t\t' + str(tie_index) + '\t\t' + str(differences))
        print('#' * 100)

        results.append((_+1, player_wins, banker_wins, ties, sum(differences) / len(differences) if len(differences) else 0, max(differences) if len(differences) else 0, min(differences) if len(differences) else 0, tie_index, differences))

        

    tie_indexes = []
    for result in results:
        for index_ in result[-2]:
            tie_indexes.append(index_)

    filename = f'tie_index_{str(t)}.csv'    
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['index'])
        for index in tie_indexes:
            writer.writerow([index])



    diffs = []
    for result in results:
        for diff_ in result[-1]:
            diffs.append(diff_)

    filename = f'tie_diff_{str(t)}.csv'
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['diff'])
        for diff in diffs:
            writer.writerow([diff])

    ## store the results to a csv file

    filename = f'baccarat_{str(t)}.csv'

    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['#','Player Wins', 'Banker Wins', 'Ties', 'Avg Diff', 'Max Diff', 'Min Diff', 'Tie Index', 'Differences'])
        for result in results:
            writer.writerow(result)