# %% imports
import pathlib
import numpy as np
import xlsxwriter

# %% Load file
data_path = pathlib.Path("./data")

with open(data_path / "wordlist-eng.txt") as file:
    word_list = file.readlines()

word_list = [word.strip() for word in word_list]
# %%
NUM_BOARDS_TO_GENERATE = 10
# %%


def generate_grid_and_answers():
    selected_words = np.random.choice(word_list, size=25, replace=False)
    red_first = np.random.choice([True, False])

    num_words_to_pick_red = 6 + int(red_first)
    num_words_to_pick_blue = 6 + int(~red_first)

    selected_indices = set()

    index_range = range(25)

    red_indices = set()
    blue_indices = set()

    num_red_selected_words = 0
    num_blue_selected_words = 0

    while num_red_selected_words < num_words_to_pick_red:
        selected_index = np.random.choice(index_range)
        if selected_index not in selected_indices:
            selected_indices.add(selected_index)

            red_indices.add(selected_index)
            num_red_selected_words += 1

    while num_blue_selected_words < num_words_to_pick_blue:
        selected_index = np.random.choice(index_range)
        if selected_index not in selected_indices:
            selected_indices.add(selected_index)

            blue_indices.add(selected_index)
            num_blue_selected_words += 1

    assassin_index = np.random.choice(
        list((set(index_range) - blue_indices - red_indices))
    )

    return selected_words, red_indices, blue_indices, assassin_index


# %%
def write_word_matrix(worksheet, selected_words, format):

    worksheet.set_column(0, 5, 25 * 0.75)
    worksheet.set_default_row(139 * 0.75)

    for col in range(5):
        for row in range(5):
            worksheet.write(row, col, selected_words[row * 5 + col], format)


def convert_index_to_row_column(index):
    col = index % 5
    row = int(index / 5)

    return row, col


def add_answers_to_worksheet(
    worksheet,
    red_indices,
    blue_indices,
    assassin_index,
    red_format,
    blue_format,
    assassin_format,
):

    for red_index in red_indices:
        row, col = convert_index_to_row_column(red_index)
        worksheet.conditional_format(
            row, col, row, col, {"type": "no_errors", "format": red_format}
        )
    for blue_index in blue_indices:
        row, col = convert_index_to_row_column(blue_index)
        worksheet.conditional_format(
            row, col, row, col, {"type": "no_errors", "format": blue_format}
        )

    row, col = convert_index_to_row_column(assassin_index)
    worksheet.conditional_format(
        row, col, row, col, {"type": "no_errors", "format": assassin_format}
    )


workbook = xlsxwriter.Workbook("./output/codenames.xlsx")

for i in range(NUM_BOARDS_TO_GENERATE):
    selected_words, red_indices, blue_indices, assassin_index = (
        generate_grid_and_answers()
    )

    worksheet = workbook.add_worksheet(f"Grid {i}")
    answer_worksheet = workbook.add_worksheet(f"Answers {i}")
    format = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 16,
        }
    )
    blue_format = workbook.add_format({"bg_color": "#3464eb"})
    red_format = workbook.add_format({"bg_color": "#eb4034"})
    assassin_format = workbook.add_format({"bg_color": "#5e595e"})

    write_word_matrix(worksheet, selected_words, format)
    write_word_matrix(answer_worksheet, selected_words, format)

    add_answers_to_worksheet(
        answer_worksheet,
        red_indices,
        blue_indices,
        assassin_index,
        red_format,
        blue_format,
        assassin_format,
    )


workbook.close()
# %%
