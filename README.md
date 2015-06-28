# EnergyExportDatabrowser

An open-source reboot of the existing Energy Export Databrowser.

The [Energy Export Databrowser](http://mazamascience.com/OilExport/) was introduced in June of 2008 in a [post on The Oil Drum](http://www.theoildrum.com/node/4127).

The original databrowser used data from the [British Petroleum Statistical Review](http://www.bp.com/en/global/corporate/about-bp/energy-economics/statistical-review-of-world-energy.html) to provide graphics that emphasize the multi-decadal trends in energy production and consumption. Here is the motivation given in the original databrowser:

> Access to fossil fuels is one of the most important issues of our time. The world's largest economies are extremely
> dependent upon imported supplies of oil and gas. Understanding who produces and consumes oil, coal and natural gas is
> critical today and will remain so in the years ahead.

The importance of access to fossil fuels has not dimished in the years since and the Energy Export Databrowser is ready for a complete overhaul. This will be an opportunity for us to exercise our new skills in interactive data visualization and open source collaboration.

----

This repository consists of two subdirectories:

* StaticData/ -- python scripts to convert the BP .xlsx workbook into a set of csv files
* Databrowser/ -- Mazama Science Databrowser code to provide an interactive interface

For now, only the StaticData/ directory has code. The python files in that directory have been used to convert
recent versions of the BP Statistical Review but need to be updated to accommodate the June, 2015 release. 
Mazama Science will not have the resources to update the conversion code until late in 2015 so any contributions
would be greatly appreciated.
