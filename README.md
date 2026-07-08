# Smart CSV Analyzer

**Live Demo:** [你的 GitHub Pages 链接](https://wy-data-30.github.io/csv-data-analyzer/)

[中文说明](./README.zh-CN.md)

Smart CSV Analyzer is a browser-based CSV analysis tool for quick exploration of structured CSV files. After a user uploads a CSV file, the application performs data preview, field type detection, data quality checks, descriptive statistics, chart generation, custom grouping analysis, and basic insight generation entirely in the browser.

The project is built as a static website and does not require a backend service, database, account system, or build step.

## Features

- Upload and parse CSV files in the browser.
- Preview the first 10 rows of the dataset.
- Automatically identify numeric, categorical, date, and ID fields.
- Generate a dataset overview, including row count, field count, field type counts, missing values, and duplicate rows.
- Produce data quality reports for missing values, missing rates, duplicate rows, and numeric outliers.
- Calculate descriptive statistics for numeric fields, including mean, median, minimum, maximum, and standard deviation.
- Generate frequency statistics for categorical fields, including Top 10 categories and proportions.
- Render charts for numeric distributions, categorical Top 10 values, and date trends.
- Support custom analysis by selecting a numeric metric, a categorical grouping field, and an optional date field.
- Provide scenario templates for general data, sales data, student scores, used product prices, surveys, and user behavior logs.
- Generate cautious, data-based insights without assuming external business context.
- Support responsive layouts for desktop and mobile screens.

## Tech Stack

- HTML5
- CSS3
- Vanilla JavaScript
- [PapaParse](https://www.papaparse.com/) for CSV parsing
- [Chart.js](https://www.chartjs.org/) for data visualization

External libraries are loaded through CDN links in `index.html`.

## Project Structure

```text
.
├── index.html
├── style.css
├── script.js
├── sample-data.csv
├── sample-sales.csv
├── sample-students.csv
├── sample-used-products.csv
├── sample-survey.csv
├── sample-user-behavior.csv
└── README.md
```

## Getting Started

You can open `index.html` directly in a browser.

For the best local experience, start a static server from the project directory:

```bash
python -m http.server 8000
```

Then visit:

```text
http://localhost:8000
```

Using a local static server is recommended because some browsers restrict `fetch()` when opening files directly from the local filesystem. This can affect the "Load Sample Data" button.

## Usage

1. Open the application in a browser.
2. Upload a CSV file by selecting a file or dragging it into the upload area.
3. Review the generated data overview, schema detection, quality report, statistics, charts, and insights.
4. Use the field selectors to choose a metric, grouping field, and optional date field for custom analysis.
5. Select a scenario template and map fields manually when the dataset fits a supported analysis scenario.

The application does not require fixed column names. Template analysis is based on the fields selected by the user.

## Sample Data

The repository includes sample CSV files for testing different analysis scenarios:

- `sample-data.csv`: general order-style data for testing the full analysis workflow.
- `sample-sales.csv`: sales data for testing amount, unit price, quantity, category, region, and date trend analysis.
- `sample-students.csv`: student score data for testing score distribution, class/subject grouping, and exam date trends.
- `sample-used-products.csv`: used product price data for testing price distribution, platform, city, condition, and price outliers.
- `sample-survey.csv`: survey data for testing satisfaction, age, gender, city, and choice frequency analysis.
- `sample-user-behavior.csv`: user behavior data for testing user IDs, events, channels, devices, and event time trends.

The sample files use UTF-8 BOM encoding and include Chinese field names and values. They also contain a small number of missing values, duplicate rows, and outliers to help verify the data quality checks.

## Supported Data Types

- Numeric fields: values such as amount, quantity, score, price, duration, and rate.
- Categorical fields: values such as region, category, city, channel, device, gender, and subject.
- Date fields: values that can be parsed by the browser as dates, such as `2025-01-01` or `2025-06-01 08:00:00`.
- ID fields: identifiers such as order IDs, user IDs, student IDs, event IDs, codes, and primary keys.

## Field Detection Rules

For each column, the application calculates non-empty count, unique count, uniqueness ratio, numeric match ratio, and date match ratio. Field types are detected with the following heuristic rules:

- ID field: the field name contains keywords such as `id`, `uuid`, `code`, `key`, `number`, `no`, `编号`, `编码`, `订单号`, `账号`, `工号`, or `主键`; or the column has a very high uniqueness ratio and is not primarily numeric or date-like.
- Date field: at least 80% of non-empty values can be parsed as dates.
- Numeric field: at least 80% of non-empty values can be parsed as numbers.
- Categorical field: used as the default type when none of the rules above apply.

Field detection is heuristic. Users can override the analysis direction through custom field selectors and scenario template mappings.

## Data Quality Rules

- Missing values: empty strings, `null`, `undefined`, and whitespace-only values are treated as missing.
- Missing rate: calculated as missing values divided by total row count for each column.
- Duplicate rows: detected by comparing normalized values across all fields in each row.
- Numeric outliers: detected with the IQR method. Values below `Q1 - 1.5 * IQR` or above `Q3 + 1.5 * IQR` are marked as outliers.
- Empty CSV handling: files without headers or valid data rows show a clear error message and do not keep stale analysis results on screen.

## Charts

- Numeric distribution chart: displays numeric values as a histogram.
- Categorical Top 10 chart: displays the 10 most frequent categories as a bar chart.
- Date trend chart: aggregates record counts or numeric values by date and displays a line chart.
- Custom grouping chart: groups a selected numeric metric by a selected categorical field.
- Scenario template charts: render a primary chart, secondary chart, and optional trend chart based on the selected template and field mapping.

## Project Highlights

- Static, browser-only architecture suitable for GitHub Pages and other static hosting platforms.
- No fixed business schema; CSV files with different structures can be analyzed through automatic detection and manual field mapping.
- Privacy-friendly processing because uploaded files are parsed locally in the browser.
- Practical coverage of the common CSV exploration workflow: preview, schema detection, quality checks, statistics, visualization, and insights.
- Built-in scenario templates for common structured datasets.
- Responsive interface with scrollable tables and adaptive chart layouts.
- Supports Chinese field names and values in CSV files.

## Privacy

All CSV parsing and analysis run locally in the browser. The application does not upload CSV data to a server, store user files, or send dataset contents to a third-party API.

The page loads PapaParse and Chart.js from CDN providers. When the page loads these external scripts, the browser may request resources from the CDN host, but uploaded CSV contents remain local to the browser runtime.

## Known Limitations

- Large CSV files are limited by browser memory and single-page runtime performance.
- Date parsing depends on the browser's `Date` implementation; non-standard date formats may not be detected accurately.
- Field type detection is based on heuristic rules and may require manual confirmation for ambiguous columns.
- The project focuses on basic exploratory analysis and does not include advanced data cleaning, forecasting, or machine learning.
- Analysis results are not persisted. Refreshing the page clears the current analysis.
- CDN dependencies require network access unless the libraries are downloaded and referenced locally.

## Roadmap

- Add chunked parsing and progress indicators for large CSV files.
- Improve encoding and delimiter detection.
- Add export options for HTML, PDF, or image reports.
- Add more chart types such as box plots, scatter plots, and stacked bar charts.
- Support filtering, sorting, renaming fields, derived fields, and simple data cleaning operations.
- Add saved analysis configurations.
- Split core analysis logic into smaller modules and add automated tests.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
