import pandas as pd
import random
import argparse
import os


def random_color():
    """
    Generate random RGB code
    """
    return random.randrange(0, 256), random.randrange(0, 256),\
        random.randrange(0, 256)


def rgb2hex(r, g, b):
    """
    Convert RGB code to HEX code
    """
    return "%02x%02x%02x" % (r, g, b)


def build_color_table(n):
    """
    Build a data frame of n random non-repetitive colors with RGB and HEX codes
    """
    df_colors = pd.DataFrame(columns=["HEX", "RGB"])
    hex_list = []
    for i in range(0, n):
        r, g, b = random_color()
        hex_code = rgb2hex(r, g, b)
        while hex_code in hex_list:
            r, g, b = random_color()
            hex_code = rgb2hex(r, g, b)
        hex_list.append(hex_code)
        add = pd.DataFrame(
            [[hex_code, str(r) + "-" + str(g) + "-" + str(b)]],
            columns=df_colors.columns)
        df_colors = df_colors.append(add)
    df_colors.reset_index(drop=True, inplace=True)
    return df_colors


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--n', default='0',
                        help="Number of colors to be generated")
    args = parser.parse_args()

    random.seed(0)
    df_colors = build_color_table(int(args.n))
    dictionaries_path = "../dictionaries"
    os.makedirs(dictionaries_path, exist_ok=True)
    random_colors_path = os.path.join(
        dictionaries_path, "random_colors_" + args.n + ".csv"
    )
    df_colors.to_csv(random_colors_path)


if __name__ == "__main__":
    main()
