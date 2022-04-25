import argparse
import os
from core.flipper import PoeFlipper


ALL_POSSIBLE_CATEGORIES = ['accessory', 'armour', 'weapon', 'jewel', 'uniqueMap']


def init_argparse() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        usage="%(prog)s [OPTION]...",
    )
    parser.add_argument(
        "-gf", "--generate-file", nargs="+",
        help="Generate tables with selected categories. List of categories: https://api.poe.watch/categories. Default armour accessory weapon",
        type=str,
    )
    parser.add_argument("-cf", "--custom-filename",
                        help="Work only in pair with --generate-file(-gf). Provides custom filename for generated file.",
                        )
    parser.add_argument("-f", "--format",
                        help="Specify file format for generated file",
                        )
    parser.add_argument('-if', '--import-file',
                        help="Use custom file. File must be in the same folder as script")
    return parser


def main():
    parser = init_argparse()
    args = parser.parse_args()

    custom_file = "item_table.xlsx"
    if args.import_file:
        if not os.path.isfile(args.import_file):
            print("{} Not found".format(args.import_file))
            raise SystemExit()
        custom_file = args.import_file
    file_format = args.format if args.format else "xlsx"
    flipper = PoeFlipper("Archnemesis", file_format, custom_filename=custom_file)

    if args.custom_filename and not args.generate_file:
        print("-cf works only in pair with --generate-table(-gt)")
        raise SystemExit()
    if args.generate_file:
        custom_filename = "None"
        for arg in args.generate_file:
            if arg not in ALL_POSSIBLE_CATEGORIES:
                print("{} category not supported. Skip..".format(arg))
                args.generate_file.remove(arg)
        if args.custom_filename:
            custom_filename = args.custom_filename[0]
        flipper.generate_items_file(args.generate_file, custom_filename)

    flipper.start()


if __name__ == "__main__":
    main()
