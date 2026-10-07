"""Assertions for preserving manually maintained workbook sheets."""


def assert_manual_sheets_preserved(case, actual_workbook, source, *, excluded_names=()):
    case.assertEqual(
        {
            name: definition
            for name, definition in actual_workbook.defined_names.items()
            if name not in excluded_names
        },
        {
            name: definition
            for name, definition in source.defined_names.items()
            if name not in excluded_names
        },
    )
    case.assertEqual(actual_workbook.loaded_theme, source.loaded_theme)
    for name in ("APOIO", "RESUMO GASTOS"):
        actual, expected = actual_workbook[name], source[name]
        case.assertEqual(actual.max_row, expected.max_row)
        case.assertEqual(actual.max_column, expected.max_column)
        case.assertEqual(
            set(map(str, actual.merged_cells)), set(map(str, expected.merged_cells))
        )
        case.assertEqual(actual.data_validations, expected.data_validations)
        case.assertEqual(dict(actual.tables), dict(expected.tables))
        case.assertEqual(
            {
                str(area.sqref): actual.conditional_formatting[area]
                for area in actual.conditional_formatting
            },
            {
                str(area.sqref): expected.conditional_formatting[area]
                for area in expected.conditional_formatting
            },
        )
        for property_name in (
            "print_area",
            "print_options",
            "page_setup",
            "page_margins",
            "sheet_properties",
            "sheet_format",
            "sheet_view",
            "protection",
        ):
            case.assertEqual(
                getattr(actual, property_name), getattr(expected, property_name)
            )
        case.assertEqual(
            {
                key: (dict(value), value._style)
                for key, value in actual.column_dimensions.items()
            },
            {
                key: (dict(value), value._style)
                for key, value in expected.column_dimensions.items()
            },
        )
        case.assertEqual(
            {
                key: (dict(value), value._style)
                for key, value in actual.row_dimensions.items()
            },
            {
                key: (dict(value), value._style)
                for key, value in expected.row_dimensions.items()
            },
        )
        case.assertEqual(dict(actual.defined_names), dict(expected.defined_names))
        for row in expected:
            for cell in row:
                copied = actual[cell.coordinate]
                if not (name == "RESUMO GASTOS" and cell.coordinate == "M2"):
                    case.assertEqual(
                        copied.value, cell.value, f"{name}!{cell.coordinate}"
                    )
                case.assertEqual(
                    copied._style, cell._style, f"{name}!{cell.coordinate}"
                )
