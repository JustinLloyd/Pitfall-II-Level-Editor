Attribute VB_Name = "Components"
Option Explicit

'Public Enum ComponentOrder
'    k_ORDER_FIRST = -1
'    k_ORDER_FLOOR
'    k_ORDER_BACKGROUND
'    k_ORDER_EXIT_LEFT
'    k_ORDER_EXIT_RIGHT
'    k_ORDER_CREATURE
'    k_ORDER_ITEM
'    k_ORDER_FEATURE_LSB
'    k_ORDER_FEATURE_MSB
'    k_ORDER_LAST
'End Enum

'Public Enum FloorTypes
'    k_FLOOR_FIRST = -1
'    k_FLOOR_NONE                                ' no floor
'    k_FLOOR_PLAT_LEFT                           ' floor is a platform on left of screen with black background
'    k_FLOOR_PLAT_RIGHT                          ' floor is a platform on right of screen with black background
'    k_FLOOR_PLAT_BOTH                           ' floor is two platforms on edges of screen with black background
'    k_FLOOR_PLAT_LEFT_WATER                     ' floor is a platform on left of screen with blue background
'    k_FLOOR_PLAT_RIGHT_WATER                    ' floor is a platform on right of screen with blue background
'    k_FLOOR_PLAT_BOTH_WATER                     ' floor is a platform on edges of screen with blue background
'    k_FLOOR_WATER                               ' floor is water
'    k_FLOOR_SOLID                               ' solid floor
'    k_FLOOR_WALKWAY                             ' floor is a walkway
'    k_FLOOR_WALKWAY_SINGLE_HOLE                 ' floor is a walkway with a hole (trapdoor) in it
'    k_FLOOR_WALKWAY_THREE_HOLES                 ' floor is a walkway with three holes in it
'    k_FLOOR_RIVER_BED                           ' floor is packed earth (like a river bed)
'    k_FLOOR_WALKWAY_HOLE_WITH_LADDER            ' floor is a walkway with a hole and a ladder
'    k_FLOOR_WALKWAY_SIX_HOLES                   ' floor is a walkway with six holes
'    k_FLOOR_LAST
'End Enum

'Public Const k_STR_OPT_FLOOR_NONE = "No floor"
'Public Const k_STR_CONST_FLOOR_NONE = "k_FLOOR_NONE"
'Public Const k_STR_OPT_FLOOR_PLAT_LEFT = "Platform Left"
'Public Const k_STR_CONST_FLOOR_PLAT_LEFT = "k_FLOOR_PLAT_LEFT"
'Public Const k_STR_OPT_FLOOR_PLAT_RIGHT = "Platform Right"
'Public Const k_STR_CONST_FLOOR_PLAT_RIGHT = "k_FLOOR_PLAT_RIGHT"
'Public Const k_STR_OPT_FLOOR_PLAT_BOTH = "Platform Both"
'Public Const k_STR_CONST_FLOOR_PLAT_BOTH = "k_FLOOR_PLAT_BOTH"
'Public Const k_STR_OPT_FLOOR_PLAT_LEFT_WATER = "Platform Left with Water"
'Public Const k_STR_CONST_FLOOR_PLAT_LEFT_WATER = "k_FLOOR_PLAT_LEFT_WATER"
'Public Const k_STR_OPT_FLOOR_PLAT_RIGHT_WATER = "Platform Right with Water"
'Public Const k_STR_CONST_FLOOR_PLAT_RIGHT_WATER = "k_FLOOR_PLAT_RIGHT_WATER"
'Public Const k_STR_OPT_FLOOR_PLAT_BOTH_WATER = "Platform Both with Water"
'Public Const k_STR_CONST_FLOOR_PLAT_BOTH_WATER = "k_FLOOR_PLAT_BOTH_WATER"
'Public Const k_STR_OPT_FLOOR_WATER = "Water"
'Public Const k_STR_CONST_FLOOR_WATER = "k_FLOOR_WATER"
'Public Const k_STR_OPT_FLOOR_SOLID = "Solid"
'Public Const k_STR_CONST_FLOOR_SOLID = "k_FLOOR_SOLID"
'Public Const k_STR_OPT_FLOOR_WALKWAY = "Walkway"
'Public Const k_STR_CONST_FLOOR_WALKWAY = "k_FLOOR_WALKWAY"
'Public Const k_STR_OPT_FLOOR_WALKWAY_SINGLE_HOLE = "Walkway, Single Hole"
'Public Const k_STR_CONST_FLOOR_WALKWAY_SINGLE_HOLE = "k_FLOOR_WALKWAY_SINGLE_HOLE"
'Public Const k_STR_OPT_FLOOR_WALKWAY_THREE_HOLES = "Walkway, Three Holes"
'Public Const k_STR_CONST_FLOOR_WALKWAY_THREE_HOLES = "k_FLOOR_WALKWAY_THREE_HOLES"
'Public Const k_STR_OPT_FLOOR_RIVER_BED = "River Bed"
'Public Const k_STR_CONST_FLOOR_RIVER_BED = "k_FLOOR_RIVER_BED"
'Public Const k_STR_OPT_FLOOR_WALKWAY_HOLE_WITH_LADDER = "Walkway, Hole with Ladder"
'Public Const k_STR_CONST_FLOOR_WALKWAY_HOLE_WITH_LADDER = "k_FLOOR_WALKWAY_HOLE_WITH_LADDER"
'Public Const k_STR_OPT_FLOOR_WALKWAY_SIX_HOLES = "Walkway, Six holes"
'Public Const k_STR_CONST_FLOOR_WALKWAY_SIX_HOLES = "k_FLOOR_WALKWAY_SIX_HOLES"

'Public Enum ExitLeftTypes
'    k_EXIT_LEFT_FIRST = -1
'    k_EXIT_LEFT_OPEN                                ' no rock face at exit
'    k_EXIT_LEFT_DARK_ROCK_BLUE                      ' dark rock on blue background
'    k_EXIT_LEFT_LIGHT_ROCK_BLACK                    ' light rock on black background
'    k_EXIT_LEFT_DARK_ROCK_BLACK                     ' dark rock on black background
'    k_EXIT_LEFT_LIGHT_ROCK_GREEN                    ' light rock on green background
'    k_EXIT_LEFT_DARK_ROCK_GREEN                     ' dark rock on green background
'    k_EXIT_LEFT_PATTERN_ROCK_BLACK                  ' pattern rock on black background
'    k_EXIT_LEFT_LAST
'End Enum

'Public Const k_STR_OPT_EXIT_LEFT_OPEN = "Open Exit"
'Public Const k_STR_CONST_EXIT_LEFT_OPEN = "k_EXIT_LEFT_OPEN"
'Public Const k_STR_OPT_EXIT_LEFT_DARK_ROCK_BLUE = "Dark Rock, Blue Background"
'Public Const k_STR_CONST_EXIT_LEFT_DARK_ROCK_BLUE = "k_EXIT_LEFT_DARK_ROCK_BLUE"
'Public Const k_STR_OPT_EXIT_LEFT_LIGHT_ROCK_BLACK = "Light Rock, Black Background"
'Public Const k_STR_CONST_EXIT_LEFT_LIGHT_ROCK_BLACK = "k_EXIT_LEFT_LIGHT_ROCK_BLACK"
'Public Const k_STR_OPT_EXIT_LEFT_DARK_ROCK_BLACK = "Dark Rock, Black Background"
'Public Const k_STR_CONST_EXIT_LEFT_DARK_ROCK_BLACK = "k_EXIT_LEFT_DARK_ROCK_BLACK"
'Public Const k_STR_OPT_EXIT_LEFT_LIGHT_ROCK_GREEN = "Light Rock, Green Background"
'Public Const k_STR_CONST_EXIT_LEFT_LIGHT_ROCK_GREEN = "k_EXIT_LEFT_LIGHT_ROCK_GREEN"
'Public Const k_STR_OPT_EXIT_LEFT_DARK_ROCK_GREEN = "Dark Rock, Green Background"
'Public Const k_STR_CONST_EXIT_LEFT_DARK_ROCK_GREEN = "k_EXIT_LEFT_DARK_ROCK_GREEN"
'Public Const k_STR_OPT_EXIT_LEFT_PATTERN_ROCK_BLACK = "Pattern Rock, Black Background"
'Public Const k_STR_CONST_EXIT_LEFT_PATTERN_ROCK_BLACK = "k_EXIT_LEFT_PATTERN_ROCK_BLACK"

'Public Enum ExitRightTypes
'    k_EXIT_RIGHT_FIRST = -1
'    k_EXIT_RIGHT_OPEN                                ' no rock face at exit
'    k_EXIT_RIGHT_DARK_ROCK_BLUE                     ' dark rock on blue background
'    k_EXIT_RIGHT_LIGHT_ROCK_BLACK                   ' light rock on black background
'    k_EXIT_RIGHT_DARK_ROCK_BLACK                    ' dark rock on black background
'    k_EXIT_RIGHT_LIGHT_ROCK_GREEN                   ' light rock on green background
'    k_EXIT_RIGHT_DARK_ROCK_GREEN                    ' dark rock on green background
'    k_EXIT_RIGHT_PATTERN_ROCK_BLACK                 ' pattern rock on black background
'    k_EXIT_RIGHT_LAST
'End Enum

'Public Const k_STR_OPT_EXIT_RIGHT_OPEN = "Open Exit"
'Public Const k_STR_CONST_EXIT_RIGHT_OPEN = "k_EXIT_RIGHT_OPEN"
'Public Const k_STR_OPT_EXIT_RIGHT_DARK_ROCK_BLUE = "Dark Rock, Blue Background"
'Public Const k_STR_CONST_EXIT_RIGHT_DARK_ROCK_BLUE = "k_EXIT_RIGHT_DARK_ROCK_BLUE"
'Public Const k_STR_OPT_EXIT_RIGHT_LIGHT_ROCK_BLACK = "Light Rock, Black Background"
'Public Const k_STR_CONST_EXIT_RIGHT_LIGHT_ROCK_BLACK = "k_EXIT_RIGHT_LIGHT_ROCK_BLACK"
'Public Const k_STR_OPT_EXIT_RIGHT_DARK_ROCK_BLACK = "Dark Rock, Black Background"
'Public Const k_STR_CONST_EXIT_RIGHT_DARK_ROCK_BLACK = "k_EXIT_RIGHT_DARK_ROCK_BLACK"
'Public Const k_STR_OPT_EXIT_RIGHT_LIGHT_ROCK_GREEN = "Light Rock, Green Background"
'Public Const k_STR_CONST_EXIT_RIGHT_LIGHT_ROCK_GREEN = "k_EXIT_RIGHT_LIGHT_ROCK_GREEN"
'Public Const k_STR_OPT_EXIT_RIGHT_DARK_ROCK_GREEN = "Dark Rock, Green Background"
'Public Const k_STR_CONST_EXIT_RIGHT_DARK_ROCK_GREEN = "k_EXIT_RIGHT_DARK_ROCK_GREEN"
'Public Const k_STR_OPT_EXIT_RIGHT_PATTERN_ROCK_BLACK = "Pattern Rock, Black Background"
'Public Const k_STR_CONST_EXIT_RIGHT_PATTERN_ROCK_BLACK = "k_EXIT_RIGHT_PATTERN_ROCK_BLACK"

'Public Enum BackgroundTypes
'    k_BACKGROUND_FIRST = -1
'    k_BACKGROUND_NONE
'    k_BACKGROUND_TREES
'    k_BACKGROUND_TREE_TOPS
'    k_BACKGROUND_WATER
'    k_BACKGROUND_EARTH
'    k_BACKGROUND_LAST
'End Enum

'Public Const k_STR_OPT_BACKGROUND_NONE = "No background"
'Public Const k_STR_CONST_BACKGROUND_NONE = "k_BACKGROUND_NONE"
'Public Const k_STR_OPT_BACKGROUND_TREES = "Trees"
'Public Const k_STR_CONST_BACKGROUND_TREES = "k_BACKGROUND_TREES"
'Public Const k_STR_OPT_BACKGROUND_TREE_TOPS = "Tree Tops"
'Public Const k_STR_CONST_BACKGROUND_TREE_TOPS = "k_BACKGROUND_TREE_TOPS"
'Public Const k_STR_OPT_BACKGROUND_WATER = "Water"
'Public Const k_STR_CONST_BACKGROUND_WATER = "k_BACKGROUND_WATER"
'Public Const k_STR_OPT_BACKGROUND_EARTH = "Earth"
'Public Const k_STR_CONST_BACKGROUND_EARTH = "k_BACKGROUND_EARTH"

'Public Enum FeatureTypes
'    k_FEATURE_FIRST = -1
'    k_FEATURE_NONE
'    k_FEATURE_SAVE_POINT
'    k_FEATURE_LADDER
'    k_FEATURE_BALLOON
'    k_FEATURE_WATERFALL
'    k_FEATURE_LARA
'    k_FEATURE_VINE
'    k_FEATURE_LAST
'End Enum
'
'Public Const k_STR_OPT_FEATURE_SAVE_POINT = "Save Point"
'Public Const k_STR_CONST_FEATURE_SAVE_POINT = "k_FEATURE_SAVE_POINT"
'Public Const k_STR_CONST_FEATURE_SAVE_POINT_MSK = "k_FEATURE_SAVE_POINT_MSK"
'Public Const k_STR_OPT_FEATURE_LADDER = "Ladder"
'Public Const k_STR_CONST_FEATURE_LADDER = "k_FEATURE_LADDER"
'Public Const k_STR_CONST_FEATURE_LADDER_MSK = "k_FEATURE_LADDER_MSK"
'Public Const k_STR_OPT_FEATURE_BALLOON = "Balloon"
'Public Const k_STR_CONST_FEATURE_BALLOON = "k_FEATURE_BALLOON"
'Public Const k_STR_CONST_FEATURE_BALLOON_MSK = "k_FEATURE_BALLOON_MSK"
'Public Const k_STR_OPT_FEATURE_WATERFALL = "Waterfall"
'Public Const k_STR_CONST_FEATURE_WATERFALL = "k_FEATURE_WATERFALL"
'Public Const k_STR_CONST_FEATURE_WATERFALL_MSK = "k_FEATURE_WATERFALL_MSK"
'Public Const k_STR_OPT_FEATURE_LARA = "Lara"
'Public Const k_STR_CONST_FEATURE_LARA = "k_FEATURE_LARA"
'Public Const k_STR_CONST_FEATURE_LARA_MSK = "k_FEATURE_LARA_MSK"
'Public Const k_STR_OPT_FEATURE_VINE = "Vine"
'Public Const k_STR_CONST_FEATURE_VINE = "k_FEATURE_VINE"
'Public Const k_STR_CONST_FEATURE_VINE_MSK = "k_FEATURE_VINE_MSK"
'
'Public Enum CreatureTypes
'    k_CREATURE_FIRST = -1
'    k_CREATURE_NONE
'    k_CREATURE_BAT
'    k_CREATURE_CONDOR
'    k_CREATURE_EEL
'    k_CREATURE_FROG
'    k_CREATURE_SCORPION
'    k_CREATURE_RABID_BAT
'    k_CREATURE_FIRE_ANT
'    k_CREATURE_PIRANHA
'    k_CREATURE_WALK_FROG
'    k_CREATURE_SNAKE
'    k_CREATURE_TREE_SNAKE
'    k_CREATURE_CROCODILE
'    k_CREATURE_LAST
'End Enum
'
'Public Const k_STR_OPT_CREATURE_NONE = "No Creature"
'Public Const k_STR_CONST_CREATURE_NONE = "k_CREATURE_NONE"
'Public Const k_STR_OPT_CREATURE_BAT = "Bat"
'Public Const k_STR_CONST_CREATURE_BAT = "k_CREATURE_BAT"
'Public Const k_STR_OPT_CREATURE_CONDOR = "Condor"
'Public Const k_STR_CONST_CREATURE_CONDOR = "k_CREATURE_CONDOR"
'Public Const k_STR_OPT_CREATURE_EEL = "Electric Eel"
'Public Const k_STR_CONST_CREATURE_EEL = "k_CREATURE_EEL"
'Public Const k_STR_OPT_CREATURE_FROG = "Hole Frog"
'Public Const k_STR_CONST_CREATURE_FROG = "k_CREATURE_HOLE_FROG"
'Public Const k_STR_OPT_CREATURE_SCORPION = "Scorpion"
'Public Const k_STR_CONST_CREATURE_SCORPION = "k_CREATURE_SCORPION"
'Public Const k_STR_OPT_CREATURE_RABID_BAT = "Rabid Bat"
'Public Const k_STR_CONST_CREATURE_RABID_BAT = "k_CREATURE_RABID_BAT"
'Public Const k_STR_OPT_CREATURE_FIRE_ANT = "Fire Ant"
'Public Const k_STR_CONST_CREATURE_FIRE_ANT = "k_CREATURE_FIRE_ANT"
'Public Const k_STR_OPT_CREATURE_PIRANHA = "Piranha"
'Public Const k_STR_CONST_CREATURE_PIRANHA = "k_CREATURE_PIRANHA"
'Public Const k_STR_OPT_CREATURE_WALk_FROG = "Walking Frog"
'Public Const k_STR_CONST_CREATURE_WALk_FROG = "k_CREATURE_WALk_FROG"
'Public Const k_STR_OPT_CREATURE_SNAKE = "Coiled Snake"
'Public Const k_STR_CONST_CREATURE_SNAKE = "k_CREATURE_SNAKE"
'Public Const k_STR_OPT_CREATURE_TREE_SNAKE = "Tree Snake"
'Public Const k_STR_CONST_CREATURE_TREE_SNAKE = "k_CREATURE_TREE_SNAKE"
'Public Const k_STR_OPT_CREATURE_CROCODILE = "Crocodile"
'Public Const k_STR_CONST_CREATURE_CROCODILE = "k_CREATURE_CROCODILE"
'
'Public Enum ItemTypes
'    k_ITEM_FIRST = -1
'    k_ITEM_NONE
'    k_ITEM_STONE_RAT
'    k_ITEM_QUICKCLAW_CAT
'    k_ITEM_DIAMOND_RING
'    k_ITEM_RHONDA_GIRL
'    k_ITEM_GOLD_BAR_LEFT
'    k_ITEM_GOLD_BAR_RIGHT
'    k_ITEM_HAT
'    k_ITEM_LARA
'    k_ITEM_LAMP
'    k_ITEM_ROPE
'    k_ITEM_LAST
'End Enum
'
'Public Const k_STR_OPT_ITEM_NONE = "No Item"
'Public Const k_STR_CONST_ITEM_NONE = "k_ITEM_NONE"
'Public Const k_STR_OPT_ITEM_STONE_RAT = "Stone Rat"
'Public Const k_STR_CONST_ITEM_STONE_RAT = "k_ITEM_STONE_RAT"
'Public Const k_STR_OPT_ITEM_QUICKCLAW_CAT = "Quickclaw Cat"
'Public Const k_STR_CONST_ITEM_QUICKCLAW_CAT = "k_ITEM_QUICKCLAW"
'Public Const k_STR_OPT_ITEM_DIAMOND_RING = "Diamond Ring"
'Public Const k_STR_CONST_ITEM_DIAMOND_RING = "k_ITEM_DIAMOND_RING"
'Public Const k_STR_OPT_ITEM_RHONDA_GIRL = "Rhonda"
'Public Const k_STR_CONST_ITEM_RHONDA_GIRL = "k_ITEM_RHONDA_GIRL"
'Public Const k_STR_OPT_ITEM_GOLD_BAR_LEFT = "Gold Bar Left"
'Public Const k_STR_OPT_ITEM_GOLD_BAR_RIGHT = "Gold Bar Right"
'Public Const k_STR_CONST_ITEM_GOLD_BAR_LEFT = "k_ITEM_GOLD_BAR_LEFT"
'Public Const k_STR_CONST_ITEM_GOLD_BAR_RIGHT = "k_ITEM_GOLD_BAR_RIGHT"
'Public Const k_STR_OPT_ITEM_HAT = "Hat"
'Public Const k_STR_CONST_ITEM_HAT = "k_ITEM_HAT"
'Public Const k_STR_OPT_ITEM_LARA = "Lara"
'Public Const k_STR_CONST_ITEM_LARA = "k_ITEM_LARA"
'Public Const k_STR_OPT_ITEM_LAMP = "Lamp"
'Public Const k_STR_CONST_ITEM_LAMP = "k_ITEM_LAMP"
'Public Const k_STR_OPT_ITEM_ROPE = "Rope"
'Public Const k_STR_CONST_ITEM_ROPE = "k_ITEM_ROPE"
'
'
'' vcs components
'Public Enum LowNibbleTypes
'    k_LOW_FIRST = -1
'    k_LOW_NONE
'    k_LOW_WATER
'    k_LOW_EARTH
'    k_LOW_TREE_TOPS_1
'    k_LOW_TREES_1
'    k_LOW_FLOOR_TWO_HOLES_AND_LADDER
'    k_LOW_CORRUPT_1
'    k_LOW_CORRUPT_2
'    k_LOW_EARTH_FLAT_FLOOR
'    k_LOW_WALKWAY
'    k_LOW_SINGLE_HOLE
'    k_LOW_SINGLE_HOLE_AND_LADDER
'    k_LOW_RIVER
'    k_LOW_TREE_TOPS_2
'    k_LOW_TREES_2
'    k_LOW_CORRUPT_3
'    k_LOW_LAST
'End Enum

'Public Const k_STR_OPT_LOW_NONE = "Black background"
'Public Const k_STR_OPT_LOW_WATER = "Water"
'Public Const k_STR_OPT_LOW_EARTH = "Earth"
'Public Const k_STR_OPT_LOW_TREETOPS_1 = "Treetops #1"
'Public Const k_STR_OPT_LOW_TREES_1 = "Tress #1"
'Public Const k_STR_OPT_LOW_FLOOR_TWO_HOLES_AND_LADDER = "Two holes & ladder"
'Public Const k_STR_OPT_LOW_EARTH_FLAT_FLOOR = "Earth and Floor"
'Public Const k_STR_OPT_LOW_WALKWAY = "Walkway"
'Public Const k_STR_OPT_LOW_SINGLE_HOLE = "Single hole"
'Public Const k_STR_OPT_LOW_SINGLE_HOLE_AND_LADDER = "Single Hole & Ladder"
'Public Const k_STR_OPT_LOW_RIVER = "River"
'Public Const k_STR_OPT_LOW_TREE_TOPS_2 = "Tree Tops #2"
'Public Const k_STR_OPT_LOW_TREES_2 = "Trees #2"
'Public Const k_STR_OPT_LOW_CORRUPT = "Corrupt"


'Public Enum HighNibbleTypes
'    k_HIGH_FIRST = -1
'    k_HIGH_NONE
'    k_HIGH_SAVE_POINT
'    k_HIGH_BIG_HOLE_1
'    k_HIGH_QUICKCLAW
'    k_HIGH_SCORPION
'    k_HIGH_BAT
'    k_HIGH_CONDOR
'    k_HIGH_GOLD_BAR_LEFT
'    k_HIGH_STONE_RAT
'    k_HIGH_WATERFALL
'    k_HIGH_BIG_HOLE_2
'    k_HIGH_RHONDA
'    k_HIGH_DIAMOND_RING
'    k_HIGH_BALLOON
'    k_HIGH_FROG
'    k_HIGH_GOLD_BAR_RIGHT
'    k_HIGH_LAST
'End Enum

'Public Const k_STR_OPT_HIGH_NONE = "None"
'Public Const k_STR_OPT_HIGH_SAVE_POINT = "Save Point"
'Public Const k_STR_OPT_HIGH_BIG_HOLE_1 = "Big Hole #1"
'Public Const k_STR_OPT_HIGH_QUICKCLAW = "Quickclaw"
'Public Const k_STR_OPT_HIGH_SCORPION = "Scorpion"
'Public Const k_STR_OPT_HIGH_BAT = "Bat or Eel"
'Public Const k_STR_OPT_HIGH_CONDOR = "Condor"
'Public Const k_STR_OPT_HIGH_GOLD_BAR_LEFT = "Gold Bar Left"
'Public Const k_STR_OPT_HIGH_STONE_RAT = "Stone Rat"
'Public Const k_STR_OPT_HIGH_WATERFALL = "Waterfall"
'Public Const k_STR_OPT_HIGH_BIG_HOLE_2 = "Big Hole #2"
'Public Const k_STR_OPT_HIGH_RHONDA = "Rhonda"
'Public Const k_STR_OPT_HIGH_DIAMOND_RING = "Diamond Ring"
'Public Const k_STR_OPT_HIGH_BALLOON = "Balloon"
'Public Const k_STR_OPT_HIGH_FROG = "Frog"
'Public Const k_STR_OPT_HIGH_GOLD_BAR_RIGHT = "Gold Bar Right"

