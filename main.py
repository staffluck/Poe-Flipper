import argparse


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




if __name__ == "__main__":
    main()