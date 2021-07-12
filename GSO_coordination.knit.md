---
title: "GSO Coordination"
author: "Nicholas Manna"
date: '2020-12-10'
runtime: shiny
params: 
  excel_data: NA
  datapath: NA
  sheet: NA
output: html_document
knit: (function(inputFile, encoding) {rmarkdown::render(inputFile, encoding = encoding, output_file = file.path(dirname(inputFile), paste0(lubridate::today(), '_GSO_Coordination.html')))})
---

<style type="text/css">
.main-container {
  max-width: 1800px;
  margin-left: auto;
  margin-right: auto;
}
</style>















