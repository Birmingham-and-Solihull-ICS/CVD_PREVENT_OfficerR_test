# BSOL CVD Prevent Data
## Tesing production of PowerPoint slides using OfficeR

This repository contains initial exploratory code for using the officer R package.
Once it was established, it was then built into a loop for each CVDPrevent indicator to rpoduce slides.

Process involved:

+ Using ICB pptx template to cut out relevant slide masters.
+ Use 'Select' pane in Powerpoint to find labels of the different sections to populate.
+ Generate a visualisation in R, transform to a raster graphic using the related `rvg` package, insert into a 'picture' object.
+ Build out text for titles and brief auto-commentary elements on simple slides.
+ Run look, generating a list of slides.
+ Print list to file to generate the pptx.

This repository is dual licensed under the [Open Government v3]([https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/) & MIT. All code can outputs are subject to Crown Copyright.
