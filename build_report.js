const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,HeadingLevel,Table,TableRow,TableCell,WidthType,ShadingType,BorderStyle,AlignmentType,LevelFormat,Footer,PageNumber,TableOfContents}=require('docx');

const W=9026; // A4 with 1" margins
const H=(t,l=HeadingLevel.HEADING_1)=>new Paragraph({heading:l,children:[new TextRun(t)]});
const P=(t,o={})=>new Paragraph({spacing:{after:120},...o,children:Array.isArray(t)?t:[new TextRun(t)]});
const B=(t,lvl=0)=>new Paragraph({numbering:{reference:'b',level:lvl},spacing:{after:60},children:Array.isArray(t)?t:[new TextRun(t)]});
const N=(t)=>new Paragraph({numbering:{reference:'n',level:0},spacing:{after:60},children:Array.isArray(t)?t:[new TextRun(t)]});
const bold=(a,b)=>[new TextRun({text:a,bold:true}),new TextRun(b)];
const bd={style:BorderStyle.SINGLE,size:4,color:'BFBFBF'};
const borders={top:bd,bottom:bd,left:bd,right:bd};
function table(widths,head,rows){
  const cell=(t,w,h)=>new TableCell({width:{size:w,type:WidthType.DXA},borders,margins:{top:60,bottom:60,left:100,right:100},
    shading:h?{fill:'2E7D32',type:ShadingType.CLEAR,color:'auto'}:undefined,
    children:[new Paragraph({children:[new TextRun({text:t,bold:h,color:h?'FFFFFF':undefined,size:19})]})]});
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:widths,rows:[
    new TableRow({tableHeader:true,children:head.map((t,i)=>cell(t,widths[i],true))}),
    ...rows.map(r=>new TableRow({children:r.map((t,i)=>cell(t,widths[i],false))}))]});
}
const sp=()=>P('',{spacing:{after:80}});

const c=[];
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:1200,after:120},children:[new TextRun({text:'Week 1 Report',size:28,color:'2E7D32',bold:true})]}));
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:'Data Acquisition and Problem Scoping',size:44,bold:true})]}));
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:120},children:[new TextRun({text:'Machine Learning for Agriculture and Agribusiness',size:28})]}));
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:600},children:[new TextRun({text:'District-Level Crop Yield Prediction in India using Weather, Soil and Satellite Data',italics:true,size:24})]}));
c.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun('Name: Shivbhushan Gaikwad     |     Internship: Yuva Intern     |     Date: 29 September 2026')]}));
c.push(new Paragraph({children:[new TextRun({text:'Contents',bold:true,size:28})],spacing:{before:600,after:120},pageBreakBefore:true}));
const toc=[
[1,'1. Introduction and Objectives'],[2,'1.1 Week 1 objectives'],
[1,'2. Problem Definition'],[2,'2.1 Background and motivation'],[2,'2.2 Chosen problem statement'],[2,'2.3 Machine learning formulation'],[2,'2.4 Why this problem is a good fit'],[2,'2.5 Success criteria'],[2,'2.6 Scope boundaries'],
[1,'3. Data Requirements'],[2,'3.1 Required data characteristics'],
[1,'4. Annotated Outline of Public Data Sources'],[2,'4.1 Crop yield (target variable)'],[2,'4.2 Weather and climate'],[2,'4.3 Soil'],[2,'4.4 Satellite vegetation, land use and irrigation'],[2,'4.5 Source selection rationale'],
[1,'5. Data Collection Plan'],
[1,'6. Understanding the Dataset (Planned Exploration)'],
[1,'7. Data Quality Assessment Plan'],
[1,'8. Pseudo-Methodology for Future Weeks'],[2,'8.1 Data preparation'],[2,'8.2 Feature engineering'],[2,'8.3 Modelling'],[2,'8.4 Validation and evaluation'],[2,'8.5 Interpretation and reporting'],
[1,'9. Project Timeline'],[1,'10. Challenges and Risks'],[1,'11. Documentation and Decision Log'],[1,'12. Conclusion and Next Steps'],[2,'Appendix: Data Catalogue Template']];
toc.forEach(([l,t])=>c.push(new Paragraph({spacing:{after:l===1?80:40,before:l===1?100:0},indent:{left:l===1?0:500},children:[new TextRun({text:t,bold:l===1,size:l===1?23:21})]})));

// 1
c.push(new Paragraph({heading:HeadingLevel.HEADING_1,pageBreakBefore:true,children:[new TextRun('1. Introduction and Objectives')]}));
c.push(P('This report is the Week 1 deliverable of the internship. Its purpose is to start the project properly: choose a machine learning problem that matters in agriculture, identify what data it needs, find where that data can be obtained, and lay out a realistic plan for the weeks ahead. No models are built this week. The emphasis is on making sound, documented decisions that later work can rely on.'));
c.push(P('The work follows a simple six-stage workflow, and the sections of this report mirror it:'));
c.push(table([1000,3000,5026],['Stage','Workflow step','Where covered'],[
 ['1','Problem Definition','Section 2'],
 ['2','Identify Data Requirements','Section 3'],
 ['3','Collect Data (source research and acquisition plan)','Sections 4 and 5'],
 ['4','Understand Dataset (planned exploration)','Section 6'],
 ['5','Check Data Quality','Section 7'],
 ['6','Document Findings (methodology, timeline, risks, decision log)','Sections 8 to 11']]));
c.push(sp());
c.push(H('1.1 Week 1 objectives',HeadingLevel.HEADING_2));
c.push(B('Define one specific, measurable machine learning problem in the agriculture domain.'));
c.push(B('Research publicly available data on crop yields, weather, soil health and related parameters.'));
c.push(B('Annotate each candidate source with its publication frequency, credibility and granularity.'));
c.push(B('Draft a pseudo-methodology and a week-by-week timeline for data collection, preparation and analysis.'));
c.push(B('Record every decision and the reasoning behind it in a structured way.'));

// 2
c.push(H('2. Problem Definition'));
c.push(H('2.1 Background and motivation',HeadingLevel.HEADING_2));
c.push(P('Agriculture employs a large share of the workforce in India and is highly exposed to weather. Monsoon timing, dry spells, heat stress and soil condition all influence how much a farm produces in a given season. Early and reliable estimates of crop yield are valuable to several groups: state agriculture departments planning procurement and storage, crop insurers assessing risk, agribusinesses managing supply chains and input sales, and farmers deciding what to plant.'));
c.push(P('Traditional yield estimates rely on crop-cutting experiments and field surveys. These are accurate but slow, expensive and only available after the season has largely ended. A machine learning model that uses freely available weather, soil and satellite data could give an earlier and cheaper estimate, and could highlight districts at risk before a poor harvest is confirmed.'));
c.push(H('2.2 Chosen problem statement',HeadingLevel.HEADING_2));
c.push(new Paragraph({spacing:{after:160},indent:{left:360,right:360},border:{left:{style:BorderStyle.SINGLE,size:18,color:'2E7D32',space:8}},children:[new TextRun({text:'Predict the seasonal yield (tonnes per hectare) of major crops at district level in India, using historical yield records combined with seasonal weather, soil properties and satellite vegetation indices, and flag districts at elevated risk of a below-normal yield.',italics:true})]}));
c.push(H('2.3 Machine learning formulation',HeadingLevel.HEADING_2));
c.push(bullets=null);
c.pop();
c.push(B(bold('Primary task (regression): ','predict yield in tonnes per hectare for each district, crop and season.')));
c.push(B(bold('Secondary task (classification): ','label each district-season as normal or at-risk, where at-risk means yield falls more than a chosen threshold (for example 15 percent) below that district\'s recent multi-year average. The threshold is provisional and will be reviewed after exploring the data.')));
c.push(B(bold('Unit of observation: ','one row per district, crop and season (kharif or rabi) for each year.')));
c.push(B(bold('Target variable: ','yield = production divided by area, taken from official agricultural statistics.')));
c.push(B(bold('Starting scope: ','one or two staple crops (for example rice and wheat, or soybean and cotton if focusing on Maharashtra) before broadening. Starting narrow keeps the first pass manageable.')));
c.push(H('2.4 Why this problem is a good fit',HeadingLevel.HEADING_2));
c.push(B('Official yield records exist for many years and districts, so there is enough labelled data to train and validate a model.'));
c.push(B('Weather, soil and satellite data are all freely available at national or global scale, so the project can be reproduced without paid data.'));
c.push(B('The result has clear practical value, which makes success easy to explain to non-technical stakeholders.'));
c.push(B('It involves real challenges (spatial and temporal dependence, missing data, small sample sizes) that make it a meaningful learning exercise.'));
c.push(H('2.5 Success criteria',HeadingLevel.HEADING_2));
c.push(B('The model should beat two simple baselines: the district\'s previous-year yield and its multi-year average.'));
c.push(B('Errors should be reported with RMSE, MAE and R-squared for the regression task, and precision, recall and F1 for the at-risk classification.'));
c.push(B('Evaluation must use a temporal split (train on earlier years, test on later years) so results reflect real forecasting conditions rather than leaked information.'));
c.push(H('2.6 Scope boundaries',HeadingLevel.HEADING_2));
c.push(P('In scope: district-level seasonal yield, public data only, tabular and gridded inputs aggregated to district. Out of scope for this project: farm-level prediction, price forecasting, pest and disease detection, and real-time in-season deployment. These can be listed as future extensions.'));

// 3
c.push(H('3. Data Requirements'));
c.push(P('Working backward from the problem statement, the following variable groups are needed. Each group is tied to an agronomic reason so that feature choices can be justified later.'));
c.push(table([1900,3400,3726],['Variable group','Example variables','Why it matters'],[
 ['Target: crop yield','Area (ha), production (tonnes), yield (t/ha) by district, crop, season, year','The quantity to be predicted'],
 ['Rainfall','Seasonal total, monthly totals, number of rainy days, longest dry spell, monsoon onset timing','Water availability is the dominant driver in rainfed regions'],
 ['Temperature','Mean, minimum and maximum; heat stress days above crop-specific thresholds; growing degree days','Heat during flowering and grain filling reduces yield'],
 ['Other weather','Relative humidity, solar radiation, wind, evapotranspiration','Affects crop water demand and disease pressure'],
 ['Soil','Organic carbon, pH, nitrogen, phosphorus, potassium, texture (sand, silt, clay), water holding capacity','Determines nutrient supply and water retention'],
 ['Vegetation (satellite)','NDVI or EVI time series, peak greenness, season length','Direct signal of crop vigour during the season'],
 ['Irrigation and inputs','Irrigated area share, fertiliser use where available, cropping intensity','Separates irrigated from rainfed performance'],
 ['Administrative and static','District code, state, latitude, longitude, elevation, agro-climatic zone','Allows spatial joins and captures regional effects']]));
c.push(sp());
c.push(H('3.1 Required data characteristics',HeadingLevel.HEADING_2));
c.push(B(bold('Time span: ','ideally 15 to 25 years of history, so that several drought and good-monsoon years are represented.')));
c.push(B(bold('Spatial resolution: ','district-level for yield; weather, soil and satellite data are gridded and will be aggregated to district boundaries.')));
c.push(B(bold('Temporal resolution: ','yield is annual per season; weather and vegetation should be daily, weekly or monthly so that within-season features can be built.')));
c.push(B(bold('Consistent identifiers: ','district names and boundaries have changed over time in India (new districts carved out), so a mapping table will be needed to join data reliably.')));

// 4
c.push(H('4. Annotated Outline of Public Data Sources'));
c.push(P('The table below lists candidate sources. Publication frequency, credibility and granularity are recorded for each, as the task requires. Access details, licences and current coverage must be re-checked on each source\'s website before download, since portals change over time.'));
c.push(H('4.1 Crop yield (target variable)',HeadingLevel.HEADING_2));
c.push(table([1900,1500,1500,1900,2226],['Source','Publication frequency','Credibility','Granularity','Notes'],[
 ['Ministry of Agriculture and Farmers Welfare, Directorate of Economics and Statistics (Agricultural Statistics at a Glance and area, production, yield tables)','Annual, released with a lag of one to two years','High: official government statistics','State and district, by crop and season','Primary source for the target. Check for revisions to older years.'],
 ['Open Government Data Platform India (data.gov.in)','Varies by dataset','High for government-published datasets','District level for many crop datasets','Convenient CSV downloads and APIs. Quality and completeness vary by dataset.'],
 ['ICRISAT district-level database','Periodic updates','High: international research institute','District level, long historical series','Useful for a long, harmonised series with fixed district boundaries.'],
 ['FAOSTAT','Annual','High: UN agency','Country level only','Useful for context and sanity checks, not for district modelling.']]));
c.push(sp());
c.push(H('4.2 Weather and climate',HeadingLevel.HEADING_2));
c.push(table([1900,1500,1500,1900,2226],['Source','Publication frequency','Credibility','Granularity','Notes'],[
 ['India Meteorological Department (IMD) gridded rainfall and temperature','Daily, updated regularly','High: national meteorological authority','About 0.25 degree grid for rainfall, 1 degree for temperature, daily','Best local fit for India. Download procedures may need a Python helper library.'],
 ['NASA POWER','Daily, near real time','High: NASA-derived reanalysis and satellite products','About 0.5 degree grid, daily; also point queries by latitude and longitude','Easy API. Provides radiation, humidity and wind that IMD grids lack.'],
 ['ERA5 or ERA5-Land (Copernicus Climate Data Store)','Hourly, with a short delay','High: ECMWF reanalysis','About 0.1 to 0.25 degree, hourly','Rich variable set. Requires a free account and larger downloads.'],
 ['CHIRPS rainfall','Daily and monthly','High: widely used in research','About 0.05 degree','A good cross-check for rainfall against IMD.']]));
c.push(sp());
c.push(H('4.3 Soil',HeadingLevel.HEADING_2));
c.push(table([1900,1500,1500,1900,2226],['Source','Publication frequency','Credibility','Granularity','Notes'],[
 ['SoilGrids (ISRIC)','Periodic version releases','High: peer-reviewed global product','250 m grid, several depth layers','Modelled rather than measured, so uncertainty is higher locally. Good for consistent coverage.'],
 ['Soil Health Card portal (Government of India)','Cycles of a few years','High: government, measured samples','Farm and village level, aggregated to district','Measured nutrient data, but coverage and access for bulk download must be checked.'],
 ['ICAR and NBSS&LUP soil maps','Static, infrequent','High: national research institutions','Soil map units at regional scale','Useful as a static context layer and for validation.'],
 ['Harmonized World Soil Database (FAO/IIASA)','Static','High','About 1 km','A fallback if SoilGrids is impractical.']]));
c.push(sp());
c.push(H('4.4 Satellite vegetation, land use and irrigation',HeadingLevel.HEADING_2));
c.push(table([1900,1500,1500,1900,2226],['Source','Publication frequency','Credibility','Granularity','Notes'],[
 ['MODIS NDVI and EVI (NASA)','16-day composites, plus daily products','High: long-running, well validated','250 m to 1 km','Consistent since 2000. Good for district-level seasonal vegetation profiles.'],
 ['Sentinel-2 (ESA/Copernicus)','Every 5 days or so from 2015','High','10 to 20 m','Finer detail but heavier processing. Consider as a later extension.'],
 ['Land use and land cover products (for example ISRO Bhuvan, ESA WorldCover)','Annual or occasional','High','10 m to 100 m','Used to mask non-cropland pixels before averaging NDVI.'],
 ['Administrative boundaries (for example Survey of India, GADM, or community shapefiles)','Occasional','Varies: verify boundary vintage','District polygons','Essential for spatial aggregation. Must match the year of the yield data.']]));
c.push(sp());
c.push(H('4.5 Source selection rationale',HeadingLevel.HEADING_2));
c.push(P('The initial working combination is: official district yield statistics (target), IMD or NASA POWER for weather, SoilGrids for soil, and MODIS NDVI for vegetation. This mixes authoritative national data with globally consistent gridded products, all of which are free. Alternatives are kept in reserve in case any source proves hard to access or too incomplete.'));

// 5
c.push(H('5. Data Collection Plan'));
c.push(N(bold('Set up the workspace. ','Create a project folder with separate subfolders for raw data, processed data, notebooks and documentation. Raw files are never edited by hand.')));
c.push(N(bold('Download the target data first. ','Obtain district-wise area, production and yield for the chosen crops and seasons. Record the download date, URL and file version.')));
c.push(N(bold('Build a district master table. ','Create a lookup of district codes, names (with spelling variants) and boundary vintage so all other datasets can be joined to it.')));
c.push(N(bold('Acquire weather data. ','Pull daily rainfall and temperature for the full study period. Where gridded, aggregate to district by averaging the cells within each district polygon.')));
c.push(N(bold('Acquire soil data. ','Extract soil layers for each district (mean over cropland where a land-cover mask is available) at the topsoil depth most relevant to roots.')));
c.push(N(bold('Acquire satellite data. ','Download NDVI time series for the crop season windows and aggregate to district.')));
c.push(N(bold('Log everything. ','Maintain a data catalogue (a simple table) with source, licence, date accessed, variables, coverage, and known issues.')));
c.push(sp());
c.push(P([new TextRun({text:'Reproducibility: ',bold:true}),new TextRun('all downloads and aggregation steps will be scripted (Python with pandas, geopandas, xarray or rasterio) so the dataset can be rebuilt from scratch, rather than assembled by manual clicking.')]));

// 6
c.push(H('6. Understanding the Dataset (Planned Exploration)'));
c.push(P('Once data is collected, the first analytical step is to understand its structure before any modelling. The planned checks are:'));
c.push(B(bold('Structure: ','number of rows, columns, districts, crops, seasons and years; data types; the panel shape (districts by years).')));
c.push(B(bold('Target distribution: ','histogram and summary statistics of yield per crop; identify skew and unusually high or low values.')));
c.push(B(bold('Trends over time: ','yield usually rises over the years because of technology and improved varieties. This trend must be separated from weather-driven variation, for example by detrending or including a year variable.')));
c.push(B(bold('Spatial patterns: ','maps of average yield and yield variability by district to reveal regional structure.')));
c.push(B(bold('Relationships: ','correlations and scatter plots between yield and rainfall, temperature, NDVI and soil properties; check for non-linear effects such as yield falling with both very low and very high rainfall.')));
c.push(B(bold('Multicollinearity: ','many weather variables are strongly related, so redundant features will be identified early.')));
c.push(B(bold('Extreme years: ','identify known drought and flood years and confirm the data reflects them, as a sanity check.')));

// 7
c.push(H('7. Data Quality Assessment Plan'));
c.push(table([2000,3400,3626],['Issue','How it will be detected','Planned handling'],[
 ['Missing values','Count and map missingness by district, year and variable','Drop features with very high missingness. Impute the rest with district-level or neighbouring-district statistics, and add a missing-flag if informative.'],
 ['Outliers and impossible values','Range checks (for example negative rainfall or yield above plausible crop maximums), z-scores, comparison with neighbours','Investigate each case. Correct obvious unit or entry errors, otherwise flag or remove and document.'],
 ['Inconsistent units or definitions','Compare summary statistics across sources and years; read data dictionaries','Convert to consistent units (t/ha, mm, degrees Celsius) and record conversions.'],
 ['Changing district boundaries and names','Match district lists across datasets and years','Use a master mapping table. Merge or split districts consistently, or restrict to districts with stable boundaries.'],
 ['Spatial mismatch','Compare grid resolution to district size; check for districts smaller than one grid cell','Use area-weighted aggregation and note the loss of detail for small districts.'],
 ['Temporal misalignment','Confirm season definitions (kharif and rabi months) across datasets','Aggregate all inputs over the same crop-season windows.'],
 ['Duplicates','Check duplicated district, crop, season, year keys','Remove exact duplicates and resolve conflicts by source priority.'],
 ['Bias and representativeness','Check whether missing data is concentrated in particular regions or years','Report limitations and avoid claims for regions poorly covered.'],
 ['Data leakage risk','Review each feature for information not available at prediction time','Exclude leaking features and use time-based splits.']]));

// 8
c.push(H('8. Pseudo-Methodology for Future Weeks'));
c.push(P('The steps below outline the intended approach. Details will be refined as the data is understood.'));
c.push(H('8.1 Data preparation',HeadingLevel.HEADING_2));
c.push(P('FOR each source: download, store raw copy, log metadata\nThe following steps are then applied:',{}) );
c.pop();
c.push(B('For each source: download, store an untouched raw copy, and log metadata.'));
c.push(B('Clean and harmonise: standardise units, names and codes; apply the district mapping table.'));
c.push(B('Aggregate gridded data to district polygons for each crop-season window.'));
c.push(B('Merge everything into one analysis table keyed by district, crop, season and year.'));
c.push(H('8.2 Feature engineering',HeadingLevel.HEADING_2));
c.push(B('Season-window weather features: total rainfall, mean and extreme temperature, heat-stress days, dry-spell length.'));
c.push(B('Stage-based features: weather during sowing, vegetative growth, flowering and grain filling, since sensitivity differs by stage.'));
c.push(B('Vegetation features: peak NDVI, area under the NDVI curve, date of peak.'));
c.push(B('Lag features: previous-year yield and rolling multi-year district average.'));
c.push(H('8.3 Modelling',HeadingLevel.HEADING_2));
c.push(B('Baselines: previous-year yield, district average, and linear or ridge regression.'));
c.push(B('Main models: random forest and gradient boosting (for example XGBoost or LightGBM), which handle non-linear relationships and mixed variable types well on tabular data.'));
c.push(B('Optional extension: sequence models or spatial models if time permits and the simpler models show clear limits.'));
c.push(H('8.4 Validation and evaluation',HeadingLevel.HEADING_2));
c.push(B('Temporal split: for example train on earlier years and test on the most recent years. Also try leave-one-year-out cross-validation.'));
c.push(B('Spatial check: hold out entire districts or states to test generalisation to unseen regions.'));
c.push(B('Metrics: RMSE, MAE, R-squared; precision, recall and F1 for at-risk detection; error broken down by state, crop and year type (normal, drought).'));
c.push(H('8.5 Interpretation and reporting',HeadingLevel.HEADING_2));
c.push(B('Feature importance and partial dependence or SHAP values to explain which factors drive predictions.'));
c.push(B('Final report with clear limitations and recommendations.'));

// 9
c.push(H('9. Project Timeline'));
c.push(P('A realistic timeline, assuming an eight-week internship. It should be adjusted to the actual programme length.'));
c.push(table([1100,2600,3400,1926],['Week','Focus','Key activities','Output'],[
 ['1','Problem scoping and data research','Define problem, survey data sources, plan collection, write this report','Project plan (this document)'],
 ['2','Data acquisition','Download yield, weather, soil and satellite data; build district master table; data catalogue','Raw data and catalogue'],
 ['3','Cleaning and integration','Resolve missing data, units and boundary issues; aggregate to district; merge','Clean analysis dataset'],
 ['4','Exploratory analysis','Distributions, trends, maps, correlations, quality report','EDA notebook and findings'],
 ['5','Feature engineering and baselines','Build features; train baseline models; set up validation scheme','Baseline results'],
 ['6','Modelling','Train and tune main models; compare against baselines','Model comparison'],
 ['7','Evaluation and interpretation','Error analysis, spatial checks, feature importance, at-risk classification','Evaluation report'],
 ['8','Final reporting','Consolidate documentation, visuals, limitations and future work','Final report and presentation']]));
c.push(sp());
c.push(P('A buffer of a few days is built into weeks 2 and 3, because data access and cleaning are the most common sources of delay in projects of this kind.'));

// 10
c.push(H('10. Challenges and Risks'));
c.push(table([2300,3400,3326],['Challenge','Why it matters','Mitigation'],[
 ['Data access and download friction','Some portals lack bulk download or APIs, and file formats vary','Prioritise sources with APIs. Keep alternates in reserve. Start downloads early in Week 2.'],
 ['Reporting lag and revisions in yield data','Latest years may be missing or provisional','Use complete years for training and note provisional years separately.'],
 ['Small sample size per district','Only one yield value per district per season per year, so limited data','Pool districts in one model, use regional features, and prefer simpler models with regularisation.'],
 ['Spatial and temporal autocorrelation','Neighbouring districts and consecutive years are not independent, which inflates naive validation scores','Use temporal and spatial hold-out validation.'],
 ['Technology trend in yields','Rising yields can be mistaken for weather effects','Detrend or include a time variable, and evaluate on anomalies too.'],
 ['Irrigation masks weather effects','Irrigated districts respond less to rainfall','Include irrigated area share and consider modelling irrigated and rainfed groups separately.'],
 ['Cloud cover in satellite data','Monsoon-season clouds corrupt optical NDVI','Use composited products and quality flags, and smooth the time series.'],
 ['Coarse grids for small districts','Loss of local detail','Use area-weighting and acknowledge this limitation.'],
 ['Unobserved factors','Pests, diseases, management, seed variety and market conditions are not in the data','State limitations clearly and avoid over-claiming causal explanations.'],
 ['Time and compute limits','Large raster downloads can be slow','Restrict to needed regions and years and process in batches.']]));

// 11
c.push(H('11. Documentation and Decision Log'));
c.push(P('Every important decision is recorded with the alternatives considered and the reasoning, so that the work can be reviewed and reproduced. The initial Week 1 entries are shown below, and the log will be updated each week.'));
c.push(table([600,2100,2100,2000,2226],['No.','Decision','Alternatives considered','Reason','Status'],[
 ['1','Problem: district-level seasonal crop yield prediction','Price forecasting; pest or disease detection; farm-level yield','Data is free and available in bulk, output is useful, and scope is achievable in the time','Confirmed'],
 ['2','Geographic scope: India, district level','Single state; global country-level','District data is the best balance of detail and availability. May narrow to one state if data proves hard to collect.','Provisional'],
 ['3','Start with one or two staple crops','All major crops at once','Keeps the first pass manageable','Provisional'],
 ['4','Target: yield in t/ha, plus at-risk flag','Production in tonnes only','Yield removes the effect of changing cultivated area','Confirmed'],
 ['5','Validation: temporal and spatial splits','Random train-test split','Random splits leak information across years and neighbouring districts','Confirmed'],
 ['6','Use public data only','Paid or proprietary data','Reproducibility and cost','Confirmed'],
 ['7','Model families: linear baselines, random forest, gradient boosting','Deep learning first','Strong performance on small tabular data and easier to interpret','Provisional']]));

// 12
c.push(H('12. Conclusion and Next Steps'));
c.push(P('Week 1 has defined a specific machine learning problem, mapped the data required to solve it, identified credible public sources for each variable group, and set out a methodology, timeline and risk register. The main open questions are practical ones: which sources can actually be downloaded in bulk with acceptable coverage, and how cleanly district boundaries can be matched across datasets over time.'));
c.push(P('Immediate next steps for Week 2:'));
c.push(B('Confirm access to district yield data and finalise the crop and state scope.'));
c.push(B('Download a small sample from each weather, soil and satellite source to test formats and aggregation before full-scale collection.'));
c.push(B('Create the district master mapping table and the data catalogue.'));
c.push(sp());
c.push(H('Appendix: Data Catalogue Template',HeadingLevel.HEADING_2));
c.push(table([1500,1500,1500,1500,1500,1526],['Dataset','Source and URL','Date accessed','Variables','Coverage (years, area)','Known issues']  ,[['','','','','',''],['','','','','','']]));

const doc=new Document({
 styles:{default:{document:{run:{font:'Calibri',size:22}}},
  paragraphStyles:[
   {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:32,bold:true,color:'2E7D32',font:'Calibri'},paragraph:{spacing:{before:320,after:160},outlineLevel:0}},
   {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,color:'1B5E20',font:'Calibri'},paragraph:{spacing:{before:220,after:100},outlineLevel:1}}]},
 numbering:{config:[
  {reference:'b',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]},
  {reference:'n',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
 sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1440,bottom:1440,left:1440,right:1440}}},
  children:c}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync(require("path").join(__dirname,"..","docs","Week1_Data_Acquisition_and_Problem_Scoping.docx"),b);console.log('ok')});
