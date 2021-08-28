import argparse
import os
from flipper import PoeFlipper

def init_argparse() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        usage="%(prog)s [OPTION]...",
    )
    parser.add_argument(
        "-gt", "--generate-table", nargs="+",
        help="Generate tables with selected categories. List of categories: https://api.poe.watch/categories. Default armour accessory weapon",
        type=str,
    )
    parser.add_argument("-cf", "--custom-filename",
                        help="Work only in pair with --generate-table(-gt). Provides custom filename for generated table.",
                        )
    parser.add_argument('-if', '--import-file',
                        help="Use custom excel table. File must be in the same folder as script")
    return parser


def main():
    parser = init_argparse()
    args = parser.parse_args()

    custom_file = "item_table.xlsx"
    if args.import_file:
        if not os.path.isfile(args.import_file):
            print("{} Not found".format(args.import_file))
            return 0
        custom_file = args.import_file
    flipper = PoeFlipper("Expedition", custom_file)

    if args.custom_filename and not args.generate_table:
        print("-cf works only in pair with --generate-table(-gt)")
        return 0
    if args.generate_table:
        custom_filename = "None"
        for arg in args.generate_table: 
            if arg not in ALL_POSSIBLE_CATEGORIES:
                print("{} category not supported. Skip..".format(arg))
                args.generate_table.remove(arg)
        if args.custom_filename:
            custom_filename = args.custom_filename[0]
        flipper.generate_items_table(args.generate_table, custom_filename)

    flipper.start()


if __name__ == "__main__":
    main()