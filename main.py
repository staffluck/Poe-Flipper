import argparse
import requests
import xlsxwriter


class PoeFlipper:

    def __init__(self, league, categories_to_flip=['armour', 'accessory', 'weapon']):
        self.league = league
        self.categories_to_flip = categories_to_flip
        self.poewatch_get_url = "https://api.poe.watch/get?category={}&league={league}".format("{}", league=league)

    def generate_items_table(self):
        workbook = xlsxwriter.Workbook("item_table.xlsx")
        ws = workbook.add_worksheet()
        row = 1
        for category in self.categories_to_flip:
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
                if item.get('explicits'):
                    for i in item['explicits']:
                        explicits += "{} & ".format(i)
                print(item['category'], row)
                ws.write(row, column, item['category'])
                ws.write(row, column+1, item['group'])
                ws.write(row, column+2, item['name'])
                ws.write(row, column+3, explicits)
                row += 1
            row += 1

        workbook.close()


def init_argparse() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        usage="%(prog)s [OPTION]...",
    )
    parser.add_argument(
        "-gt", "--generate-table", action="store_true"
    )
    return parser


def main():
    parser = init_argparse()
    args = parser.parse_args()
    flipper = PoeFlipper("Expedition")
    if args.generate_table:
        flipper.generate_items_table()


if __name__ == "__main__":
    main()