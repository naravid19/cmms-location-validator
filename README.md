<a id="readme-top"></a>

<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![Unlicense License][license-shield]][license-url]

<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/naravid19/cmms-location-validator">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">Location Validator</h3>

  <p align="center">
    A powerful tool to validate Master Data - Location before importing into CMMS.
    <br />
    <a href="https://github.com/naravid19/cmms-location-validator"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/naravid19/cmms-location-validator">View Demo</a>
    &middot;
    <a href="https://github.com/naravid19/cmms-location-validator/issues">Report Bug</a>
    &middot;
    <a href="https://github.com/naravid19/cmms-location-validator/issues">Request Feature</a>
  </p>
</div>

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->

## About The Project

[![Product Name Screen Shot][product-screenshot]](https://github.com/naravid19/cmms-location-validator)

The **Location Validator** is a specialized tool designed to ensure the integrity and accuracy of location data before it is imported into a Computerized Maintenance Management System (CMMS). It automates the validation process, checking for format consistency, duplicate entries, and adherence to specific coding standards (System, EQ, Component).

Key Features:

- **Automated Validation**: Quickly checks thousands of location entries against defined rules.
- **User-Friendly GUI**: Simple and intuitive interface for selecting files and running validations.
- **Detailed Reporting**: Generates comprehensive Excel reports highlighting errors and suggesting corrections.
- **Customizable**: Easily configure input files and database references.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### Built With

This project is built using Python and several powerful libraries to handle data processing and the graphical user interface.

- [![Python][Python.org]][Python-url]
- [![Pandas][Pandas.pydata.org]][Pandas-url]
- [![CustomTkinter][CustomTkinter]][CustomTkinter-url]
- [![OpenPyXL][OpenPyXL]][OpenPyXL-url]

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- GETTING STARTED -->

## Getting Started

To get a local copy up and running follow these simple example steps.

### Prerequisites

You need to have Python installed on your machine. You can download it from [python.org](https://www.python.org/).

### Installation

1.  Clone the repo
    ```sh
    git clone https://github.com/naravid19/cmms-location-validator.git
    ```
2.  Install Python packages
    ```sh
    pip install -r requirements.txt
    ```

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- USAGE EXAMPLES -->

## Usage

1.  **Run the Application**:
    You can run the application using Python:

    ```sh
    python gui_app.py
    ```

    Or run the executable if you have built it.

2.  **Configure Settings**:

    - **Sheet Name**: Enter the name of the sheet containing location data.
    - **Input File**: Select your Excel file with the location data to be validated.
    - **Database Code**: Select the reference database Excel file.

3.  **Run Validation**:
    Click the "Run Validation" button. The application will process the data and generate a report in the same directory as your input file.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ROADMAP -->

## Roadmap

- [x] Core Validation Logic
- [x] Graphical User Interface (GUI)
- [x] Executable Build Script
- [ ] Advanced Reporting Features
- [ ] Multi-language Support

See the [open issues](https://github.com/naravid19/cmms-location-validator/issues) for a full list of proposed features (and known issues).

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTRIBUTING -->

## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1.  Fork the Project
2.  Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3.  Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4.  Push to the Branch (`git push origin feature/AmazingFeature`)
5.  Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- LICENSE -->

## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTACT -->

## Contact

Narawit - [@naravid19](https://github.com/naravid19)

Project Link: [https://github.com/naravid19/cmms-location-validator](https://github.com/naravid19/cmms-location-validator)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ACKNOWLEDGMENTS -->

## Acknowledgments

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/)
- [Best-README-Template](https://github.com/othneildrew/Best-README-Template)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->

[contributors-shield]: https://img.shields.io/github/contributors/naravid19/cmms-location-validator.svg?style=for-the-badge
[contributors-url]: https://github.com/naravid19/cmms-location-validator/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/naravid19/cmms-location-validator.svg?style=for-the-badge
[forks-url]: https://github.com/naravid19/cmms-location-validator/network/members
[stars-shield]: https://img.shields.io/github/stars/naravid19/cmms-location-validator.svg?style=for-the-badge
[stars-url]: https://github.com/naravid19/cmms-location-validator/stargazers
[issues-shield]: https://img.shields.io/github/issues/naravid19/cmms-location-validator.svg?style=for-the-badge
[issues-url]: https://github.com/naravid19/cmms-location-validator/issues
[license-shield]: https://img.shields.io/github/license/naravid19/cmms-location-validator.svg?style=for-the-badge
[license-url]: https://github.com/naravid19/cmms-location-validator/blob/main/LICENSE
[product-screenshot]: images/screenshot.png
[Python.org]: https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white
[Python-url]: https://www.python.org/
[Pandas.pydata.org]: https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white
[Pandas-url]: https://pandas.pydata.org/
[CustomTkinter]: https://img.shields.io/badge/CustomTkinter-000000?style=for-the-badge&logo=python&logoColor=white
[CustomTkinter-url]: https://github.com/TomSchimansky/CustomTkinter
[OpenPyXL]: https://img.shields.io/badge/OpenPyXL-000000?style=for-the-badge&logo=python&logoColor=white
[OpenPyXL-url]: https://openpyxl.readthedocs.io/
