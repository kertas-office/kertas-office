"""Builds the Slides templates in src/templates/: python3 scripts/templates/make-slides.py
Twelve 16:9 decks for business, school and personal use. Each carries its own open-licence fonts
(fonts.py), so it looks the same on any computer, in Kertas and in PowerPoint; the pictures are
abstract stand-ins to replace. Previews: slide-previews.mjs, then slide-preview-pictures.py."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from kit import Deck, XL_CHART_TYPE
from layouts import (cover_mesh, cover_split, cover_type, section, agenda, statement, bento, stats, chart_insight,
                     donut, timeline, compare, quote, team, image_text, gallery, full_photo, plans, table_slide,
                     process, checklist, three_cards, closing)

OUT = os.path.join(os.path.dirname(__file__), '..', '..', 'src', 'templates')
COL = XL_CHART_TYPE.COLUMN_CLUSTERED
BAR = XL_CHART_TYPE.BAR_CLUSTERED
LINE = XL_CHART_TYPE.LINE_MARKERS


def theme(**k):
    base = dict(on_dark='FFFFFF', on_dark_muted='B8B8C7', on_accent='FFFFFF', on_accent_muted='EDEBFF', dark_line='34344A')
    base.update(k)
    return base


DECKS = {}


def deck(fn):
    DECKS[fn.__name__.replace('_', '-')] = fn
    return fn


# ------------------------------------------------------------------ business
@deck
def pitch_deck():
    t = theme(title='Pitch deck', short='Northwind  ·  Seed 2027', head='Plus Jakarta Sans', body='Plus Jakarta Sans',
              bg='FFFFFF', surface='F4F2FA', ink='0B0A12', muted='5B5870', line='E6E3EF', accent='6D28D9', accent2='F97316', accent3='06B6D4',
              dark='0E0B1F', on_dark_accent='C4B5FD', on_accent_muted='E9E2FF', mesh=['1A1033', '6D28D9', 'DB2777', '2563EB', 'F97316'],
              photo=['3B2A6B', '8B5CF6', 'F472B6', 'FDBA74'])
    d = Deck(t)
    cover_mesh(d, 'Seed round 2027', 'The order page every small shop deserves', 'Northwind turns phone, chat and walk-in orders into one simple list.', 'NORTHWIND  ·  PRESENTED BY YOUR NAME  ·  MONTH 2027')
    statement(d, 'Small shops lose ', '1 in 6 orders', ' to paper notes and missed messages.', kicker='The problem')
    three_cards(d, 'Our solution', 'One place for every order', [
        ('Share one link', 'A page with the menu or catalogue, made in ten minutes from a phone.'),
        ('See every order', 'Phone, chat and walk-in orders land in the same live list.'),
        ('Keep buyers posted', 'Customers get a message when their order is ready. No more calls.')])
    stats(d, 'Market', 'A big, under-served market', [('2.4M', 'Small shops', 'In our first three countries'), ('610K', 'Sell online today', 'Mostly through chat apps'), ('45K', 'Our 3-year goal', 'Paying shops, 7% of online sellers')], note='Replace with your own figures and their sources.')
    chart_insight(d, 'Traction', 'Growing 42% a month', COL, ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], [('Paying shops', (40, 72, 118, 190, 265, 380))], [t['accent']], ('380', 'paying shops in June', 'Word of mouth brings 6 in 10 new shops. Monthly churn is 3%.'))
    bento(d, 'Business model', 'Simple pricing that grows with the shop', ('$19', 'per shop, per month', 'One plan with everything included. 1.5% on card payments taken through Northwind.'),
          [('Payback in 4 months', 'Each shop costs us $70 to win and pays $19 a month.'), ('82% gross margin', 'Hosting and messages are our only running costs.')],
          ('Next: a plan for small chains', 'Shops with up to ten branches share one order list and stock count, from $49 a month.'))
    team(d, 'Team', 'Builders who know retail', [('Name Surname', 'CEO', 'Ran a family shop for ten years before Northwind.'), ('Name Surname', 'CTO', 'Built two payment apps used by 1M people.'), ('Name Surname', 'Sales', 'Signed the first 300 shops door to door.')])
    closing(d, 'Raising $1.5M to reach 5,000 shops', '50% product, 35% sales and marketing, 15% operations. 18 months of runway.', ['you@northwind.co', 'northwind.co', '+00 000 000 000'])
    return d


@deck
def business_review():
    t = theme(title='Business review', short='Q3 2026 business review', head='Manrope', body='Manrope',
              bg='FFFFFF', surface='F1F5F9', ink='0F172A', muted='526077', line='E2E8F0', accent='2563EB', accent2='14B8A6', accent3='F59E0B',
              dark='0B1220', on_dark_accent='93C5FD', on_accent_muted='DBEAFE', on_dark_muted='A9B4C6', dark_line='243047',
              mesh=['0B1220', '1D4ED8', '0EA5E9', '14B8A6'], photo=['1E3A8A', '3B82F6', '5EEAD4', 'E0F2FE'])
    d = Deck(t)
    cover_mesh(d, 'Quarterly business review', 'Q3 2026: our best quarter yet', 'Sales, customers and operations, against the plan.', 'TEAM NAME  ·  DATE')
    agenda(d, 'Today we cover', [('Results', 'The quarter in four numbers'), ('Revenue', 'Month by month, against plan'), ('Customers', 'Where growth came from'), ('Wins and misses', 'What worked, what did not'), ('Next quarter', 'Priorities, owners and dates'), ('Questions', 'Open discussion')])
    stats(d, 'Results at a glance', 'Ahead of plan on growth, behind on margin', [('$4.2M', 'Revenue', '8% above plan'), ('1,280', 'New customers', '12% above plan'), ('38%', 'Gross margin', '2 points below plan'), ('94%', 'On-time delivery', 'On plan')])
    chart_insight(d, 'Revenue', 'September was a record month', COL, ['Jul', 'Aug', 'Sep'], [('Plan', (1250, 1300, 1350)), ('Actual', (1310, 1390, 1500))], ['CBD5E1', t['accent']], ('+11%', 'over plan in September', 'The new reseller channel added $140K. July was held back by the warehouse move.'), labels=False)
    donut(d, 'Customers', 'Where new customers came from', [('Resellers', 0.42, 'New channel, launched in July'), ('Website', 0.33, 'Steady, up 4% on Q2'), ('Referrals', 0.17, 'From the loyalty scheme'), ('Events', 0.08, 'Two trade fairs')], [t['accent'], t['accent2'], t['accent3'], '94A3B8'], ('1,280', 'new customers'), None)
    compare(d, 'Wins and misses', 'What we learned this quarter', ('Went well', ['Reseller channel beat its target', 'Support replies within 2 hours', 'Two large renewals signed early']), ('To fix', ['Shipping costs rose 9%', 'Sales hiring is behind plan', 'Margin below plan in September']))
    table_slide(d, 'Next quarter', 'Priorities for Q4', [['Priority', 'Owner', 'Target', 'Due'], ['Bring shipping costs back to plan', 'Operations', '-6% cost per order', '15 Dec'], ['Hire three account managers', 'Sales', '3 hires', '30 Nov'], ['Launch the annual plan', 'Product', '200 sign-ups', '1 Dec'], ['Grow the reseller channel', 'Partnerships', '$500K revenue', '31 Dec']], [5.3, 2.3, 2.5, 1.8])
    closing(d, 'Thank you', 'Questions and discussion', ['Presenter name', 'team@company.com'])
    return d


@deck
def project_update():
    t = theme(title='Project update', short='Website relaunch  ·  Week 12', head='Inter', body='Inter',
              bg='FFFFFF', surface='F3F6F4', ink='0B1512', muted='51625C', line='E1E8E4', accent='059669', accent2='0EA5E9', accent3='F59E0B',
              dark='06201A', on_dark_accent='6EE7B7', on_accent_muted='D1FAE5', on_dark_muted='A7C4BA', dark_line='1D3B33',
              mesh=['06201A', '059669', '0EA5E9', '047857'], photo=['064E3B', '10B981', '7DD3FC', 'ECFDF5'])
    d = Deck(t)
    cover_type(d, 'Project update  ·  Week 12 of 20', 'Website relaunch', 'On track for a 1 March launch', 'PROJECT LEAD  ·  DATE', size=96)
    stats(d, 'Status', 'On track, with one risk to watch', [('12 / 20', 'Weeks done', 'Design signed off, first ten pages built'), ('1', 'Risk to watch', 'Content from sales is a week late'), ('58%', 'Budget used', 'At 60% of the schedule')])
    timeline(d, 'Timeline', 'Milestones', [('WEEK 1', 'Kick-off', 'Goals and team agreed'), ('WEEK 5', 'Design', 'Look and pages signed off'), ('WEEK 11', 'Build', 'Pages and checkout'), ('WEEK 15', 'Testing', 'Phones, tablets, payments'), ('WEEK 20', 'Launch', 'Go live and announce')], now=2)
    table_slide(d, 'Risks', 'Risks and what we are doing', [['Risk', 'Impact', 'What we are doing', 'Owner'], ['Late content from sales', 'High', 'Daily check-in; marketing writes drafts', 'Name'], ['Payment provider sign-off', 'Medium', 'Contract sent, call booked for Friday', 'Name'], ['Two testers away in week 16', 'Low', 'Extra tester booked from the agency', 'Name']], [3.7, 1.6, 4.9, 1.7])
    checklist(d, 'Next two weeks', 'What happens next', ['Finish the product pages', 'Send the content list to sales', 'Start testing on phones and tablets', 'Set up the payment test account'], ('Decision needed', 'Approve the launch date by Friday', 'We need a yes on 1 March so the press release and ads can be booked.'))
    closing(d, 'Questions?', 'Next update: Monday, week 14', ['Project lead', 'lead@company.com'])
    return d


@deck
def sales_proposal():
    t = theme(title='Sales proposal', short='Proposal for Client Name', head='Fraunces', body='Manrope', head_tracking=-20,
              bg='FBF8F3', surface='F2EBE0', ink='1C1917', muted='6B6157', line='E7DED1', accent='B45309', accent2='0F766E', accent3='9F1239',
              dark='1F1A16', on_dark_accent='FCD34D', on_dark_muted='C9BFB3', dark_line='3A322B', on_accent_muted='FDE9C8',
              mesh=['2A1D12', 'B45309', '9F1239', 'F59E0B'], photo=['78350F', 'D97706', 'FCD34D', 'FEF3C7'])
    d = Deck(t)
    cover_split(d, 'Proposal for Client Name', 'A calmer, faster front desk', 'How we will set up your bookings, payments and reminders in six weeks.', 'YOUR COMPANY  ·  DATE')
    statement(d, 'You told us your team spends ', '12 hours a week', ' on the phone confirming bookings.', kicker='What we heard')
    process(d, 'Our approach', 'Six weeks, four steps', [('Listen', 'Two days with your team to map how bookings work today.'), ('Set up', 'Online booking, payments and your price list.'), ('Train', 'Short sessions for the front desk, with a guide to keep.'), ('Support', 'Three months of help, then a check-in each quarter.')])
    image_text(d, 'What you get', 'Everything in one place', 'Clients book and pay online. Your team sees the day at a glance and reminders go out by themselves.', ['Online booking page in your colours', 'Card and bank payments', 'Text and email reminders', 'Reports each month'], seed=12)
    plans(d, 'Investment', 'Choose what fits', [('Essentials', '$2,400', 'one-off set-up', ['Booking page', 'Payments', 'Email reminders']), ('Complete', '$3,900', 'one-off set-up', ['Everything in Essentials', 'Text reminders', 'Staff training', '3 months of support']), ('Care plan', '$190', 'per month, optional', ['Updates and fixes', 'Quarterly check-in', 'Priority help'])])
    quote(d, 'No-shows fell by half in the first month, and the phone finally stopped ringing.', 'Name Surname', 'Owner, a client clinic')
    closing(d, 'Ready when you are', 'Sign by the 15th to start on the 1st of next month.', ['Your Name', 'you@yourcompany.com', '+00 000 000 000'])
    return d


@deck
def marketing_plan():
    t = theme(title='Marketing plan', short='Marketing plan 2027', head='Space Grotesk', body='Space Grotesk', head_tracking=-30,
              bg='FFFFFF', surface='F3F3EE', ink='111111', muted='5C5C58', line='E4E4DD', accent='FF4D2E', accent2='2B59FF', accent3='B9E62C',
              dark='111111', on_dark_accent='B9E62C', on_dark_muted='B5B5AE', dark_line='2E2E2C', on_accent_muted='FFE1DA',
              mesh=['111111', 'FF4D2E', '2B59FF', '7C3AED'], photo=['1F2937', 'FF4D2E', '2B59FF', 'F5F5F0'])
    d = Deck(t)
    cover_type(d, 'Marketing plan  ·  2027', 'Louder, for less.', 'How we grow sign-ups 3x on the same budget', 'BRAND TEAM  ·  DATE', size=110)
    stats(d, 'Where we are', '2026 in numbers', [('48K', 'Website visits a month', 'Up 20% on 2025'), ('2.1%', 'Visit to sign-up', 'Industry average is 2.5%'), ('$38', 'Cost per sign-up', 'Mostly from paid search')], dark=True)
    three_cards(d, 'Audience', 'Who we are talking to', [('Busy owners', 'Run a shop or studio, decide fast, read on their phone.'), ('Office managers', 'Compare options, need a case to show their boss.'), ('Freelancers', 'Price-aware, trust reviews and friends.')])
    donut(d, 'Budget', 'Where the money goes', [('Social video', 0.35, 'Short videos, made in-house'), ('Search', 0.25, 'Down from 45% last year'), ('Creators', 0.25, 'Ten small creators a quarter'), ('Events', 0.15, 'Two fairs and a meet-up')], [t['accent'], t['accent2'], t['accent3'], '9CA3AF'], ('$240K', 'for the year'), 'Replace with your own budget split.')
    timeline(d, 'Calendar', 'The year at a glance', [('Q1', 'New look', 'Refresh site and ads'), ('Q2', 'Creators', 'First creator wave'), ('Q3', 'Big launch', 'Product 2.0 campaign'), ('Q4', 'Holidays', 'Gift offer and recap')], now=0)
    chart_insight(d, 'Goal', 'Sign-ups per month', LINE, ['Jan', 'Mar', 'May', 'Jul', 'Sep', 'Nov'], [('2026', (900, 980, 1010, 1060, 1100, 1150)), ('2027 goal', (1000, 1400, 1900, 2400, 2900, 3400))], ['9CA3AF', t['accent']], ('3x', 'sign-ups by December', 'Video and creators replace half of our search spend, at a lower cost per sign-up.'), labels=False)
    closing(d, 'Let\'s make noise.', 'Questions, ideas and volunteers welcome.', ['Brand team', 'brand@company.com'])
    return d


# ------------------------------------------------------------------ school
@deck
def lesson():
    t = theme(title='Lesson', short='Science  ·  The water cycle', head='Outfit', body='Outfit', head_tracking=-20,
              bg='FFFDF8', surface='F3F0FF', ink='1E1B3A', muted='5E5A7A', line='E7E3F5', accent='7C3AED', accent2='F59E0B', accent3='10B981',
              dark='1E1B4B', on_dark_accent='FCD34D', on_dark_muted='C7C3E6', dark_line='34306A', on_accent_muted='EDE9FE',
              mesh=['1E1B4B', '7C3AED', '0EA5E9', 'F59E0B'], photo=['1E3A8A', '38BDF8', 'A7F3D0', 'FEF9C3'])
    d = Deck(t)
    cover_mesh(d, 'Science  ·  Year 6', 'The water cycle', 'Where rain comes from, and where it goes', 'TEACHER NAME  ·  LESSON 3 OF 6')
    checklist(d, 'Today', 'By the end of this lesson you can', ['Name the four stages of the water cycle', 'Explain what makes water evaporate', 'Draw the cycle with arrows and labels', 'Spot the cycle in everyday life'], ('Starter question', 'Where does a puddle go on a sunny day?', 'Think on your own for one minute, then tell the person next to you.'))
    process(d, 'The big idea', 'Four stages, round and round', [('Evaporation', 'The sun warms water and it rises as vapour.'), ('Condensation', 'Vapour cools high up and forms clouds.'), ('Precipitation', 'Drops get heavy and fall as rain or snow.'), ('Collection', 'Water gathers in seas, rivers and the ground.')])
    image_text(d, 'Look closer', 'Clouds are tiny drops', 'A cloud is made of millions of water drops so small they float. When they join up and grow heavy, they fall.', ['One raindrop holds about a million cloud drops', 'Cold air makes clouds form faster'], seed=14)
    statement(d, 'The water you drink today may once have been drunk by a ', 'dinosaur', '.', kicker='Fun fact', dark=True)
    three_cards(d, 'Your turn', 'Activity: draw the cycle', [('Draw', 'Sketch a hill, the sea, the sun and a cloud.'), ('Label', 'Add the four stages in the right places.'), ('Explain', 'Add arrows and one sentence for each stage.')])
    closing(d, 'Great work today!', 'Homework: find the water cycle at home and take a photo or draw it.', ['Next lesson: Weather'], use_mesh=True)
    return d


@deck
def research():
    t = theme(title='Research project', short='Research project  ·  Your Name', head='Instrument Serif', body='Inter', head_light=True, head_tracking=-10,
              bg='FFFFFF', surface='F4F4F2', ink='141414', muted='5E5E5A', line='E6E6E1', accent='1D4ED8', accent2='DC2626', accent3='16A34A',
              dark='141414', on_dark_accent='93C5FD', on_dark_muted='BDBDB6', dark_line='2D2D2A', on_accent_muted='DBEAFE',
              mesh=['0F172A', '1D4ED8', '334155', '0EA5E9'], photo=['334155', '64748B', 'CBD5E1', 'F8FAFC'])
    d = Deck(t)
    cover_type(d, 'Research project  ·  Department name', 'Does green space make us happier?', 'A survey of 400 people in three towns', 'YOUR NAME  ·  SUPERVISOR  ·  DATE', size=78)
    statement(d, 'People who live near parks report ', 'higher wellbeing', ', but is it the park, or the people who choose to live there?', kicker='Research question')
    three_cards(d, 'Method', 'How we found out', [('Survey', '400 adults in three towns answered 20 questions on wellbeing and habits.'), ('Mapping', 'We measured each home\'s distance to the nearest park.'), ('Comparison', 'We compared towns with similar incomes and ages.')])
    chart_insight(d, 'Findings', 'Wellbeing rises with park visits', COL, ['Never', 'Monthly', 'Weekly', 'Daily'], [('Wellbeing score (0-10)', (5.8, 6.4, 7.1, 7.5))], [t['accent']], ('+1.7', 'points from never to daily', 'The link holds after allowing for income and age. Distance alone mattered less than visits.'), number_format='0.0')
    table_slide(d, 'Findings', 'Results by town', [['Town', 'People', 'Live within 500 m', 'Visit weekly', 'Wellbeing'], ['Town A', '140', '62%', '48%', '7.0'], ['Town B', '132', '41%', '35%', '6.5'], ['Town C', '128', '28%', '22%', '6.1']], [3.0, 1.9, 2.5, 2.4, 2.1], note='Wellbeing on a 0 to 10 scale; averages per town.')
    compare(d, 'Discussion', 'What it means, and its limits', ('What we can say', ['Visiting parks often goes with higher wellbeing', 'The pattern is the same in all three towns', 'Visits matter more than distance']), ('What we cannot say', ['That parks cause the difference', 'That it holds for children', 'That it lasts over years']))
    closing(d, 'Thank you', 'Questions and feedback welcome', ['Your Name', 'you@university.edu'], use_mesh=True)
    return d


@deck
def workshop():
    t = theme(title='Workshop', short='Workshop  ·  Better meetings', head='Plus Jakarta Sans', body='Plus Jakarta Sans',
              bg='FFFFFF', surface='EEF8F6', ink='0B1A18', muted='4F6662', line='DCEDEA', accent='0D9488', accent2='F43F5E', accent3='F59E0B',
              dark='042F2B', on_dark_accent='5EEAD4', on_dark_muted='A3C9C3', dark_line='134E48', on_accent_muted='CCFBF1',
              mesh=['042F2B', '0D9488', 'F43F5E', '0F766E'], photo=['134E4A', '2DD4BF', 'FDA4AF', 'FEF3C7'])
    d = Deck(t)
    cover_split(d, 'Half-day workshop', 'Better meetings in half the time', 'Tools you can use in your next meeting, tomorrow morning.', 'FACILITATOR NAME  ·  DATE', seed=16)
    agenda(d, 'The plan for today', [('Warm-up', 'What makes a meeting good?'), ('Purpose first', 'One line that says why we meet'), ('Break', '15 minutes'), ('Running it', 'Roles, timeboxes and notes'), ('Practice', 'Run a 10-minute meeting'), ('Wrap-up', 'What you will try this week')], kicker='Agenda')
    section(d, '01', 'Purpose first', 'If you cannot say why you are meeting in one line, cancel it.')
    three_cards(d, 'Three questions', 'Before you send the invite', [('Why meet?', 'Write the purpose in one line. Put it in the invite.'), ('Who must come?', 'Only people who decide or do the work. Others get notes.'), ('What comes out?', 'A decision, a plan or a list of owners. Say which.')])
    process(d, 'Practice', 'Run a 10-minute meeting', [('Pick roles', 'Chair, timekeeper and note-taker.'), ('Set the line', 'Agree the purpose out loud.'), ('Timebox', 'Three topics, three minutes each.'), ('Close', 'Read back decisions and owners.')])
    quote(d, 'A meeting without a purpose is just a calendar event that costs money.', 'Name Surname', 'Facilitator')
    closing(d, 'Try one thing this week', 'Pick one tool from today and use it in your next meeting.', ['Facilitator name', 'hello@company.com'])
    return d


# ------------------------------------------------------------------ personal and events
@deck
def portfolio():
    t = theme(title='Portfolio', short='Your Name  ·  Portfolio', head='DM Serif Display', body='Manrope', head_light=True, head_tracking=-10,
              bg='F6F5F1', surface='ECEAE3', ink='18181B', muted='5F5E58', line='DEDBD2', accent='C2410C', accent2='18181B', accent3='A16207',
              dark='18181B', on_dark_accent='FDBA74', on_dark_muted='B9B7B0', dark_line='2E2E33', on_accent_muted='FFEDD5',
              mesh=['18181B', '7C2D12', 'C2410C', 'A8A29E'], photo=['44403C', 'A8A29E', 'FDBA74', 'F5F5F4'])
    d = Deck(t)
    cover_split(d, 'Portfolio 2027', 'Your Name, designer', 'Brands, packaging and spaces for independent makers.', 'YOURNAME.COM  ·  CITY', seed=18, size=60)
    statement(d, 'I help small brands look ', 'as good as they are', '.', kicker='About me')
    gallery(d, 'Selected work', 'Recent projects', [('Bakery rebrand', 'Identity, packaging'), ('Coffee bar', 'Interior, signage'), ('Ceramics shop', 'Website, photos'), ('Book fair', 'Posters, wayfinding')])
    full_photo(d, 'Case study: a bakery rebrand', 'New name, logo and boxes for a family bakery. Sales grew 30% in a year.', seed=44)
    process(d, 'How I work', 'From first call to launch', [('Listen', 'A morning in your shop to learn the story.'), ('Sketch', 'Three directions, rough and quick.'), ('Refine', 'One direction, made complete.'), ('Hand over', 'Files, guides and a launch kit.')])
    quote(d, 'She understood what we were about before we could say it ourselves.', 'Name Surname', 'Founder, a bakery client', use_mesh=False)
    closing(d, 'Let\'s work together', 'Booking projects from March.', ['hello@yourname.com', 'yourname.com', '@yourname'], use_mesh=False)
    return d


@deck
def event():
    t = theme(title='Event', short='Annual Gala 2027', head='Fraunces', body='Outfit', head_tracking=-20,
              bg='FFFFFF', surface='F7F3EC', ink='16130F', muted='6B6359', line='E9E1D5', accent='B08D57', accent2='7F1D1D', accent3='1F2937',
              dark='121014', on_dark_accent='E9C98B', on_dark_muted='BFB6AA', dark_line='2C2830', on_accent_muted='FBF1DD',
              mesh=['121014', '3B1E2B', 'B08D57', '1E293B'], photo=['1C1917', '7F1D1D', 'E9C98B', 'FAF5EB'])
    d = Deck(t)
    cover_mesh(d, 'You are invited', 'The Annual Gala 2027', 'An evening of dinner, music and giving for the city library.', 'SATURDAY 14 JUNE  ·  7 PM  ·  THE GRAND HALL', seed=21)
    stats(d, 'The evening', 'At a glance', [('7 pm', 'Doors open', 'Drinks in the garden'), ('8 pm', 'Dinner', 'Three courses, seated'), ('10 pm', 'Dancing', 'Live band until late')], dark=True)
    timeline(d, 'Programme', 'How the night unfolds', [('7:00', 'Welcome', 'Drinks and music'), ('8:00', 'Dinner', 'Three courses'), ('9:15', 'Auction', 'Ten lots, all for the library'), ('10:00', 'Dancing', 'Live band'), ('12:00', 'Close', 'Carriages')], now=-1)
    image_text(d, 'The venue', 'The Grand Hall', 'A 1920s hall in the old town, with a garden for drinks. Step-free entry and parking close by.', ['12 Example Street, City', 'Dress code: black tie optional', 'Dietary needs: tell us when you reply'], seed=24)
    quote(d, 'Last year you gave 4,000 children a library card. This year, let\'s build them a reading room.', 'Name Surname', 'Head of the library', seed=23)
    closing(d, 'Please reply by 1 May', 'Tickets $120 each, or $1,100 for a table of ten.', ['events@library.org', '+00 000 000 000'], seed=21)
    return d


@deck
def photo_story():
    t = theme(title='Photo story', short='A week in Lisbon', head='Instrument Serif', body='Inter', head_light=True, head_tracking=-10,
              bg='FFFFFF', surface='F3F2EF', ink='111111', muted='5E5C57', line='E6E4DF', accent='0E7490', accent2='E11D48', accent3='CA8A04',
              dark='111111', on_dark_accent='67E8F9', on_dark_muted='BDBAB3', dark_line='2B2B2B', on_accent_muted='CFFAFE',
              mesh=['0C1A22', '0E7490', 'F59E0B', '155E75'], photo=['0C4A6E', '38BDF8', 'FDBA74', 'FEF3C7'])
    d = Deck(t)
    full_photo(d, 'A week in Lisbon', 'Seven days, five neighbourhoods, too many custard tarts.', seed=50)
    image_text(d, 'Day 1', 'Up the hills of Alfama', 'We got lost on purpose: narrow lanes, washing lines and the sound of fado from open windows.', seed=51, left=True)
    gallery(d, 'Days 2 to 4', 'Small moments', [('Tram 28', 'Early, before the crowds'), ('Market', 'Fish, fruit and flowers'), ('Viewpoint', 'Sunset over the river')], seed=52)
    full_photo(d, 'Day 5: the coast', 'A train to the sea, cold water and the best grilled sardines.', seed=56)
    statement(d, 'Best thing we ate: ', 'custard tarts', ', warm, with cinnamon. Twice a day.', kicker='Highlight')
    closing(d, 'Obrigado, Lisbon', 'Until next time.', ['Your Name', 'Month 2027'], use_mesh=True, seed=50)
    return d


@deck
def minimal():
    t = theme(title='Minimal', short='Presentation title', head='Inter', body='Inter', head_tracking=-40,
              bg='FFFFFF', surface='F4F4F5', ink='09090B', muted='5F5F66', line='E4E4E7', accent='2563EB', accent2='09090B', accent3='A1A1AA',
              dark='09090B', on_dark_accent='93C5FD', on_dark_muted='A1A1AA', dark_line='27272A', on_accent_muted='DBEAFE',
              mesh=['09090B', '1E3A8A', '2563EB', '18181B'], photo=['27272A', '71717A', 'D4D4D8', 'FAFAFA'])
    d = Deck(t)
    cover_type(d, 'Subtitle or team name', 'Presentation title', 'One line that says what this talk is about', 'YOUR NAME  ·  DATE', size=96)
    agenda(d, 'What we will cover', [('First topic', 'A short description'), ('Second topic', 'A short description'), ('Third topic', 'A short description'), ('Next steps', 'What happens after today')])
    statement(d, 'Put the one idea people should remember ', 'right here', '.', kicker='Key message')
    three_cards(d, 'Three points', 'Three things to know', [('First point', 'One or two sentences that explain it.'), ('Second point', 'One or two sentences that explain it.'), ('Third point', 'One or two sentences that explain it.')])
    image_text(d, 'Picture and text', 'Show, then tell', 'Put a picture on the right and say in two sentences why it matters.', ['A supporting point', 'Another supporting point'], seed=60)
    compare(d, 'Compare', 'Before and after', ('Before', ['A point about how it was', 'Another point', 'A third point']), ('After', ['A point about how it is now', 'Another point', 'A third point']))
    closing(d, 'Thank you', 'Questions?', ['you@example.com'], use_mesh=False)
    return d


if __name__ == '__main__':
    only = sys.argv[1:]
    for name, make in DECKS.items():
        if only and name not in only:
            continue
        size = make().save(os.path.join(OUT, f'{name}.pptx'))
        print(f'{name}: {size // 1024} KB')
