import requests
import os
import openpyxl
from json.decoder import JSONDecodeError
from typing import List, Tuple

from core.data_providers import BaseDateProvider, XlsxDataProvider
from core.models import Item

POETRADE_HEADERS = {'User-Agent': 'agent47daun@gmail.com'}


class PoeFlipper:

    def __init__(self, league: str, data_provider: BaseDateProvider = XlsxDataProvider, filename="item_table.xlsx"):
        self.league: str = league
        self.filename: str = filename
        self.data_provider: BaseDateProvider = data_provider

        self.items: List[Item] = []

        self.poewatch_get_url = "https://api.poe.watch/get?category={}&league={league}".format("{}", league=league)
        self.poetrade_search = "https://www.pathofexile.com/api/trade/search/{}".format(league)
        self.poetrade_fetch = "https://www.pathofexile.com/api/trade/fetch/{}?query={}"
        self.poetrade_stats = "https://www.pathofexile.com/api/trade/data/stats"

    def convert_items_data(self, parsed_items_data: List[dict]) -> None:

        # Converting items stats into data like POETRADE_ITEM_ID:ROLL_RANGE
        def convert_mod(mod: str) -> Tuple:
            try:
                if mod.startswith("(") or mod.startswith("+("):
                    mod_raw = mod.split("(")[1].split(")")
                    mod_range = mod_raw[0]
                    mod_text = "#" + mod_raw[1]
                    if mod.startswith("+"):
                        mod_text = "+" + mod_text
                    mod_range = mod_range.split("-")
                elif mod.startswith("Socketed Gems are"):
                    mod_range = [mod[36:39].strip(), ]
                    mod_text = mod[:36] + " # " + mod[39:]
                    mod_text = mod_text.replace("  ", " ")
                elif mod.count("(") == 2:
                    # How it works:
                    # mod = Adds (65-75) to (100-110) Physical Damage
                    # mod.split("(") => ['Adds ', '65-75) to ', '100-110) Physical Damage']
                    # mod.split("(")[1].split(")") => ('65-75', ' to ')
                    # mod.split("(")[2].split(")") => ('100-110', ' Physical Damage')
                    # mod.split("(")[0] + "#" + to + "#" + bonus => Adds # to # Physical Damage
                    mod_range_first, to = mod.split("(")[1].split(")")
                    mod_range_second, bonus = mod.split("(")[2].split(")")
                    mod_range = [mod_range_first.split("-"), mod_range_second.split("-")]
                    mod_text = mod.split("(")[0] + "#" + to + "#" + bonus
                else:
                    mod_raw = mod.split("(")[1].split(")")
                    mod_range = mod_raw[0]
                    mod_text = "#" + mod_raw[1]
            except IndexError:
                #  Handling mods without range
                return (False, False)

            try:
                if converted_stats_mods.get(mod_text):
                    mod_id = converted_stats_mods[mod_text]
                elif converted_stats_mods.get(mod_text + " (Local)"):
                    mod_id = converted_stats_mods[mod_text + " (Local)"]
                else:
                    mod_id = converted_stats_mods[mod_text[1:]]
            except KeyError:
                #  Handling mods that have no POETRADE_ITEM_ID conversion
                return (False, False)

            return (mod_id, mod_range)

        try:
            stats_request = requests.get(self.poetrade_stats, headers=POETRADE_HEADERS)
            stats_data = stats_request.json()
        except JSONDecodeError:
            print("Stats fetch failed. Try again in 10 sec..")
            raise SystemExit

        converted_stats_mods = {}
        stats = stats_data['result']
        for stat in stats:
            for b in stat['entries']:
                converted_stats_mods[b['text']] = b['id']

        result_items_data = []
        for item in parsed_items_data:
            links_price = False  # Price depends on links(6-links)
            explicits_converted = {}  # Price depends on explicit rolls
            implicits_converted = {}  # Price depends on implicit rolls if item corrupted

            if item.get('explicits'):
                explicits = item['explicits'].split("&")
                for explicit in explicits:
                    explicit_id, explicit_range = convert_mod(explicit)
                    if not explicit_id:
                        continue
                    explicits_converted[explicit_id] = explicit_range

            if item.get('implicits'):
                implicits = item['implicits'].split("&")
                for implicit in implicits:
                    implicit_id, implicit_range = convert_mod(implicit)
                    if not implicit_id:
                        continue
                    implicits_converted[implicit_id] = implicit_range

            if item['group'] == "bodyarmours":
                links_price = True

            del item["explicits"]
            del item["implicits"]
            item_model = Item(explicits=explicits_converted, implicits=implicits_converted, depends_on_links=links_price, **item)
            result_items_data.append(item_model)
        return result_items_data

    def parse_file(self) -> None:
        parsed_items_data = self.data_provider.parse_file(self.filename)
        result_items_data = self.convert_items_data(parsed_items_data)
        self.items.extend(result_items_data)

    # def check_price(self, item) -> None:
    #     search_query = {
    #         "query": {
    #             "status": {
    #                 "option": "online"
    #             },
    #             "name": item["name"],

    #         }
    #     }
    #     try:
    #         request = requests.post(self.poetrade_search, data=)
    #     except JSONDecodeError as e:
    #         pass

    def start(self) -> None:
        if os.path.isfile(self.filename):
            print("Found {} table. Processing..".format(self.filename))
        else:
            print("File not found. Generating table.. ")
            self.generate_items_table(['armour', 'accessory', 'weapon'])
            return self.start()

        self.parse_file()
        for item in self.items:
            # zxc = self.check_price(item)
            print(item)

    def generate_items_table(self, categories_to_flip, custom_filename=None) -> None:
        categories = []
        for category in categories_to_flip:
            try:
                request = requests.get(self.poewatch_get_url.format(category))
                items_data = request.json()
                categories.append(items_data)
            except JSONDecodeError as e:
                print(e, request)
                continue
        self.data_provider.generate_file(self.filename, categories)
