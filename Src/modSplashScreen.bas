Attribute VB_Name = "modSplashScreen"
Option Explicit
'variables used throughout the program
Public dblbuybooks As Double
Public dblbuysupplies As Double
Public dblbuysales As Double
Public intcount As Integer

Public intcoupon As Integer
Public dblsubtotal As Double
Public dbltotal As Double
Public dblcoupondeduction As Double
Public dbltax As Double
Public intshipping As Double

Public intquantity1 As Integer
Public intquantity2 As Integer
Public intquantity3 As Integer
Public intquantity4 As Integer
Public intquantity5 As Integer
Public intquantity6 As Integer
Public intquantity7 As Integer
Public intquantity8 As Integer
Public intquantity9 As Integer
Public dbllineamount1 As Integer
Public dbllineamount2 As Integer
Public dbllineamount3 As Integer
Public dbllineamount4 As Integer
Public dbllineamount5 As Integer
Public dbllineamount6 As Integer
Public dbllineamount7 As Integer
Public dbllineamount8 As Integer
Public dbllineamount9 As Integer
'variables used in Sales Form
Public strPackdes(0 To 3) As String
Public strPackNameList(0 To 3) As String
Public dblpackpricelist(0 To 3) As Double
'variables used in Sales shopping cart form
Public strchkoutpackname(0 To 2) As String
Public dblchkoutpackprice(0 To 2) As Double
'variables used in the Supply form
Public strSupplyDes(0 To 9) As String
Public strSupplyNameList(0 To 9) As String
Public dblsupplypriceList(0 To 9) As Double
'variables used for supplies shopping cart form
Public strchkoutitemname(0 To 2) As String
Public dblchkoutitemprice(0 To 2) As String
'variables used for children's books in shopping cart form
Public strchkoutbooktitle(0 To 2) As String
Public strchkoutbooktype(0 To 2) As String
Public dblchkoutprice(0 To 2) As Double
'variables used in children's books form
Public strTitleList(0 To 9) As String
Public dblHPriceList(0 To 9) As Double
Public dblSPriceList(0 To 9) As String
Public strDescription(0 To 9) As String
'variables used in checkout form
Public strName As String
Public strCardChosen As String
Public strAddress As String
Public strCreditcardnum As String
Sub Main()
'display the splash screen and load the homepage
    frmSplash.Show
    Load frmhomepage
'sets the variable to a value of 0
    dblbuybooks = 0
    intcoupon = 0
    dblchkoutprice(0) = 0
    dblchkoutprice(1) = 0
    dblchkoutprice(2) = 0
    dblchkoutitemprice(0) = 0
    dblchkoutitemprice(1) = 0
    dblchkoutitemprice(2) = 0
    dblchkoutpackprice(0) = 0
    dblchkoutpackprice(1) = 0
    dblchkoutpackprice(2) = 0
  
    
    'strchkoutbooktitle(0) = ""
    'strchkoutbooktitle(1) = ""
    'strchkoutbooktitle(2) = ""
'book titles
    strTitleList(0) = "Harriet the Spy by Louise Fitzhugh"
    strTitleList(1) = "Nancy Drew by Carolyn Keene"
    strTitleList(2) = "Harry Potter by J.K. Rowling"
    strTitleList(3) = "The Hobbit by J.R.R. Tolkien"
    strTitleList(4) = "The Girl Who Could Fly by Victoria Forester"
    strTitleList(5) = "Inkheart by Cornelia Funke"
    strTitleList(6) = "Love You Forever by Robert Munsch"
    strTitleList(7) = "Amelia Bedelia set by Peggy Parish"
    strTitleList(8) = "A Wrinkle in Time by Madeleine L'Engle"
    strTitleList(9) = "Warriors Series by Erin Hunter"
'Hardcover book prices
    dblHPriceList(0) = "10"
    dblHPriceList(1) = "30"
    dblHPriceList(2) = "90"
    dblHPriceList(3) = "10"
    dblHPriceList(4) = "14"
    dblHPriceList(5) = "20"
    dblHPriceList(6) = "5"
    dblHPriceList(7) = "25"
    dblHPriceList(8) = "18"
    dblHPriceList(9) = "60"
'softcover  book prices
    dblSPriceList(0) = "7"
    dblSPriceList(1) = "20"
    dblSPriceList(2) = "50"
    dblSPriceList(3) = "5"
    dblSPriceList(4) = "8"
    dblSPriceList(5) = "12"
    dblSPriceList(6) = "10"
    dblSPriceList(7) = "13"
    dblSPriceList(8) = "8"
    dblSPriceList(9) = "30"
'book descriptions
    strDescription(0) = "Harriet M. Welsch is a spy. In her notebook, she writes down everything she knows about everyone, even her classmates and her best friends. Then Harriet loses track of her notebook, and it ends up in the wrong hands. Before she can stop them, her friends have read the always truthful, sometimes awful things she's written about each of them. Will Harriet find a way to put her life and her friendships back together?"
    strDescription(1) = "This specially priced starter set includes six Nancy Drew favorites: The Secret of the Old Clock, The Hidden Staircase, The Bungalow Mystery, The Mystery at Lilac Inn, The Secret of Shadow Ranch and The Secret of Red Gate Farm."
    strDescription(2) = "All seven bestselling Harry Potter titles are included in this boxed set: Harry Potter and the Philosopher's Stone, Harry Potter and the Chamber of Secrets, Harry Potter and the Prisoner of Azkaban, Harry Potter and the Goblet of Fire, Harry Potter and the Order of the Phoenix, Harry Potter and the Half-Blood Prince, Harry Potter and the Deathly Hallows."
    strDescription(3) = "The Hobbit is a tale of high adventure, undertaken by a company of dwarves in search of dragon-guarded gold. A reluctant partner in this perilous quest is Bilbo Baggins, a comfort-loving unambitious hobbit, who surprises even himself by his resourcefulness and skill as a burglar. Encounters with trolls, goblins, dwarves, elves and giant spiders, conversations with the dragon, Smaug, and a rather unwilling presence at the Battle of Five Armies are just some of the adventures that befall Bilbo."
    strDescription(4) = "You just can't keep a good girl down . . . unless you use the proper methods. Piper McCloud can fly. Just like that. Easy as pie. Sure, she hasn't mastered reverse propulsion and her turns are kind of sloppy, but she's real good at loop-the-loops. Problem is, the good folk of Lowland County are afraid of Piper. And her ma's at her wit's end. So it seems only fitting that she leave her parents' farm to attend a top-secret, maximum-security school for kids with exceptional abilities."
    strDescription(5) = "Meggie lives a quiet life alone with her father, a book-binder. But her father has a deep secret-- he possesses an extraordinary magical power. One day a mysterious stranger arrives who seems linked to her father's past. Who is this sinister character and what does he want? Suddenly Meggie is involved in a breathless game of escape and intrigue as her father's life is put in danger. Will she be able to save him in time?"
    strDescription(6) = "A young woman holds her newborn son. And looks at him lovingly. Softly she sings to him: I'll love you forever, I'll like you for always, as long as I'm living, my baby you'll be. So begins the story that has touched the hearts of millions worldwide."
    strDescription(7) = "Amelia Bedelia is the world's most literal-minded housekeeper, who causes quite a ruckus whenever she's given a chance. In Amelia Bedelia and the Surprise Shower, she arrives with a garden hose and the party is turned into an uproarious mess. In Play Ball, Amelia Bedelia, her literal-mindedness adds a new dimension to the game of baseball, and Thank You, Amelia Bedelia features Amelia Bedelia 'pairing' the vegetables and separating the eggs. In Come Back, Amelia Bedelia, Amelia Bedelia tries her hand at a variety of new jobs after Mrs. Rogers fires her for her muddles."
    strDescription(8) = "It was a dark and stormy night; Meg Murry, her small brother Charles Wallace, and her mother had come down to the kitchen for a midnight snack when they were upset by the arrival of a most disturbing stranger. A tesseract (in case the reader doesn't know) is a wrinkle in time. To tell more would rob the reader of the enjoyment of Miss L'Engle's unusual book. A Wrinkle in Time, winner of the Newbery Medal in 1963, is the story of the adventures in space and time of Meg, Charles Wallace, and Calvin O'Keefe. They are in search of Meg's father, a scientist who disappeared while engaged in secret work for the government on the tesseract problem."
    strDescription(9) = "The first story arc in the #1 nationally bestselling epic warrior cat series is now available in a beautiful box set. Contains: Into the Wild, Fire and Ice, Forest of Secrets, Rising Storm, A Dangerous Path, and The Darkest Hour."
'supply descriptions
    strSupplyDes(0) = "A set of 7 animal bookmarks that are sure to delight readers of all ages. Includes a ladybug, bee, frog, rabbit, pig, and cow bookmarks."
    strSupplyDes(1) = "These fantastic magnetic bookmarks are great for reading, clipping notes at home, or even at the office. Comes with 6 cupcake magnetic bookmarks."
    strSupplyDes(2) = "These amazing beaded bookmarks are perfect for locating the exact page you placed it in. It comes in four different styles with four differently coloured beads."
    strSupplyDes(3) = "Songbirds and butterflies amidst cheery flowers conjure a bright, sunny summer day. This hardcover journal is crafted with sturdy bookbinding, and contains lined pages of acid-free, archival paper. An inside back pocket holds notes, cards or receipts, while an elastic marker keeps your place or the journal closed. 160 pages. 5' x 7'"
    strSupplyDes(4) = "A diary is a timeless and enduring way to preserve your thoughts, creative writing, memos-to-self and so much more. This book bound diary has a beautiful cover that secures with a sturdy lock, and lined, acid-free archival pages. It makes a lovely gift. The diary comes with two keys. 192 pages. 6.25' x 8.25'"
    strSupplyDes(5) = "Ecojot's inspiring journal makes a positive impact in more ways than one: they're sustainable, B-certified, and help to deliver basic school supplies to children in need. This one has a flexible ombré blue cover with gold foil stamped quote: 'A Goal Without a Plan is Just a Wish.' - Antoine de Saint-Exupery. Printed with vegetable based ink on 100% post-consumer recycled paper. 150 lined sheets. 6' x 9'."
    strSupplyDes(6) = "Be inspired by this beautiful journal by Mari-Mi, with its filigree gold-foil origami birds, flowers and constellations printed against a rich blue ground. It has a stitched Coptic binding and creamy-smooth lined pages that are perfect for jotting down memos, notes or musings. 248 lined pages. 5.5' x 8'"
    strSupplyDes(7) = "This classic weekly planner is dated January through December and has a two-page-per-week format for a quick overview of your week. Handcrafted in Maine, it features a bonded leather cover and gilded edges that give it a luxurious look and feel. Functional features include a ribbon placeholder, a reference section, world maps, and monthly-calendars for January through December 2014."
    strSupplyDes(8) = "Show off your personality and create a keepsake of your year! The artistically designed Live Simply planner invites you to do just that-take it with you wherever you go. Spirited, unique, and inventive artwork invites creativity and moves beyond the traditional planner with pages of artwork and upbeat captions sprinkled throughout. Paste photos, ticket stubs, favorite quotes, and stickers on the blank pages in the back. With room to record your favorite moments of the year and post mementos, it's a keepsake! 4' x 6'"
    strSupplyDes(9) = "Bon vivant: Someone who lives well and enjoys fine food, drink, and schedule planning with this ooh-la-lovely Paris-inspired compact engagement calendar! Smart weekly planner format also provides space for notes and addresses. Covers 16 months (September 2013--December 2014), including the academic year. Lightweight desk engagement calendar measures 5' x 7' and fits easily in backpacks, totes, and most purses. Hardcover binding lies flat for ease of use. Handy elastic band place holder helps you stay on the right week. Vintage design is accented with gold foil and raised embossing."
'supply names
    strSupplyNameList(0) = "Animal Bookmarks"
    strSupplyNameList(1) = "Magnetic Cupcake Bookmarks"
    strSupplyNameList(2) = "Beaded Bookmarks"
    strSupplyNameList(3) = "Summer Songbird Journal"
    strSupplyNameList(4) = "Polka dot Diary"
    strSupplyNameList(5) = "Spiral Quote Journal"
    strSupplyNameList(6) = "Coptic Stitched Journal - Blue"
    strSupplyNameList(7) = "2014 Desk Agenda Key West Turquoise"
    strSupplyNameList(8) = "2014 Live Simply Planner "
    strSupplyNameList(9) = "2014 16 Month Agenda Bon Vivant"
'supply prices
    dblsupplypriceList(0) = 5
    dblsupplypriceList(1) = 10
    dblsupplypriceList(2) = 8
    dblsupplypriceList(3) = 6
    dblsupplypriceList(4) = 13
    dblsupplypriceList(5) = 14
    dblsupplypriceList(6) = 18
    dblsupplypriceList(7) = 11
    dblsupplypriceList(8) = 12
    dblsupplypriceList(9) = 6
    
'sales packs names
    strPackNameList(0) = "The Hobbit + Lord of the Rings Box Set "
    strPackNameList(1) = "Dystopian Teen Books pack"
    strPackNameList(2) = "Lorraine's Picks Contemporary pack "
    strPackNameList(3) = "Stationary pack "
'sales packs prices
    dblpackpricelist(0) = 40
    dblpackpricelist(1) = 60
    dblpackpricelist(2) = 30
    dblpackpricelist(3) = 15
'sales packs descriptions
    strPackdes(0) = "Immerse yourself in Middle-earth with Tolkien's classic masterpieces behind the films, telling the complete story of Bilbo Baggins and the Hobbits' epic encounters with Gandalf, Gollum, dragons and monsters, in the quest to destroy the One Ring. This new boxed gift set, published to celebrate the release of the first of Peter Jackson's three-part film adaptation of JRR Tolkien's The Hobbit, THE HOBBIT: AN UNEXPECTED JOURNEY, contains both titles and features cover images from both films."
    strPackdes(1) = "Includes six amazing hardcover dystopian novels: Cinder by Marissa Meyer, Angelfall by Susan Ee, Legend Series by Marie Lu, and the Darkest Minds by Alexandra Bracken. Cinder: Humans and androids crowd the raucous streets of New Beijing. A deadly plague ravages the population. From space, a ruthless lunar people watch, waiting to make their move. No one knows that Earth's fate hinges on one girl/ cyborg, Cinder.Angelfall: The Earth has been taken over by Angels who have come to destroy humans. Living among the ashes, Penryn is living day to day trying to keep her young sister and crazy mother alive. When her sister is captured by several angels,Penryn must team up with the enemy to get her beloved sister back. The Darkest Minds: In this new world, a mysterious disease has killed most of America's children, but the survivors have emerged with frightening abilities they cannot control. Ruby must try to stay alive as long as possible. " _
                    & "Legend Trilogy: The complete collection of the bestselling Legend trilogy"
    strPackdes(2) = "Contains four incredible teen novels all set in the present day. Includes: United We Spy by Ally Carter, Graffiti Moon by Cath Crowley, The Book of Broken Hearts by Sarah Ockler, and The Fault In Our Stars by John Green. United We Spy: Cammie Morgan has lost her father and her memory, but in the heart-pounding conclusion to the best-selling Gallagher Girls series, she finds her greatest mission yet. The Book of Broken Hearts: Jude has learned a lot from her older sisters, but the most important thing is this: The Vargas brothers are notorious heartbreakers. Now Jude is the only sister still living at home, and she's spending the summer helping her ailing father restore his vintage motorcycle-which means hiring a mechanic to help out. Jude must figure out how to deal with her father's illness, her sisters' oath, and trying not to fall in love. The Fault In Our Stars: Despite the tumor-shrinking medical miracle that has bought her a few years, Hazel has never been anything but terminal " _
                    & "her final chapter inscribed upon diagnosis. But when a gorgeous plot twist named Augustus Waters suddenly appears at Cancer Kid Support Group, Hazel's story is about to be completely rewritten. Graffiti Moon: Senior year is over, and Lucy has the perfect way to celebrate: tonight, she's going to find Shadow, the mysterious graffiti artist whose beautiful work appears all over the city. Instead, Lucy's suddenly stuck with Ed, on an all-night search around the city for shadow. But what Lucy can't see is the one thing that's right before her eyes."
    strPackdes(3) = "Contains the Summer Songbird Journal and the beaded bookmarks at an amazing price. Summer Sondbird Journal: Songbirds and butterflies amidst cheery flowers conjure a bright, sunny summer day. This hardcover journal is crafted with sturdy bookbinding, and contains lined pages of acid-free, archival paper. An inside back pocket holds notes, cards or receipts, while an elastic marker keeps your place or the journal closed. 160 pages. 5' x 7'. Beaded bookmarks: These amazing beaded bookmarks are perfect for locating the exact page you placed it in. It comes in four different styles with four differently coloured beads"
End Sub
