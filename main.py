import argparse
import requests
import xlsxwriter

ALL_POSSIBLE_CATEGORIES = ['accessory', 'armour', 'weapon', 'jewel', 'uniqueMap', '']


class PoeFlipper:

    def __init__(self, league):
        self.league = league
        self.poewatch_get_url = "https://api.poe.watch/get?category={}&league={league}".format("{}", league=league)

    def generate_items_table(self, categories_to_flip, custom_filename) -> None:
        workbook = xlsxwriter.Workbook("item_table.xlsx" if not custom_filename else "{}.xlsx".format(custom_filename))
        ws = workbook.add_worksheet()

        row = 1
        for category in categories_to_flip:
            try:
                request = requests.get(self.poewatch_get_url.format(category))
                items_data = request.json()
            except Exception as e:
                print(e)
                continue

            ws.write(0, 0, "Category")
            ws.write(0, 1, "Group")
            ws.write(0, 2, "Name")
            ws.write(0, 3, "Explicits")

            column = 0
            for item in items_data:
                explicits = ""
                implicits = ""
                if item.get('explicits'):
                    for i in item['explicits']:
                        explicits += "{} & ".format(i)
                if item.get('implicits'):
                    for i in item['implicits']:
                        implicits += "{} & ".format(i)

                ws.write(row, column, item['category'])
                ws.write(row, column+1, item['group'])
                ws.write(row, column+2, item['name'])
                ws.write(row, column+3, explicits)
                ws.write(row, column+4, implicits)
                row += 1
            row += 1

        workbook.close()


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

    return parser


def main():
    parser = init_argparse()
    args = parser.parse_args()
    flipper = PoeFlipper("Expedition")

    if args.custom_filename and not args.generate_table:
        print("-cf works only in pair with --generate-table(-gt)")
        return 0
    if args.generate_table:
        custom_filename = None
        for arg in args.generate_table: 
            if arg not in ALL_POSSIBLE_CATEGORIES:
                print("{} category not supported. Skip..".format(arg))
                args.generate_table.remove(arg)
        if args.custom_filename:
            custom_filename = args.custom_filename[0]
        flipper.generate_items_table(args.generate_table, custom_filename)


if __name__ == "__main__":
    main()