import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import Progress from "./Progress";
import * as Chart from 'chart.js';
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

Chart.Tooltip.positioners.event = function(elements, eventPosition) {
  return eventPosition;
};

const processRawData = (rawData) => {
  const taskColor = 'rgb(255,105,40)';
  const planningTaskColor = 'rgb(18,119,255)';
  const healthProgressingTaskColor = 'rgb(124,255,26)';
  const overDueTaskColor = 'rgb(255,8,20)';
  const maskColor = 'rgb(255, 255, 255)';
  let endDatePlanningDataSet = [];
  let startDatePlanningDataSet = [];
  let startDateDataSet = [];
  let endDateDataSet = [];
  let taskBackgroundColors = [];
  let fillLayerMaskDataSet = [];
  let fillLayerDataSet = [];
  let fillLayerBackgroundColors = [];
  let yLabels = [];
  let i;
  for (i = 0; i < rawData.length; i++) {
    const entry = rawData[i];
    const label = `${entry.customerID} ${entry.category} ${entry.revenueContribution}`;

    // push planned task
    startDatePlanningDataSet.push({x: entry.startDatePlanning, y: i});
    endDatePlanningDataSet.push({x: entry.endDatePlanning, y: i});

    // push label
    yLabels.push(label);

    if (entry.startDate !== undefined) {
      startDateDataSet.push({x: entry.startDate, y: i});
      // setup fill layer and fill layer mask
      if (entry.startDate.getTime() > entry.startDatePlanning.getTime()) {
        // start date is later than planned, need fill
        fillLayerMaskDataSet.push({x: entry.startDatePlanning, y: i});
        fillLayerDataSet.push({
          x: entry.startDate.getTime() > entry.endDatePlanning.getTime() ?
            entry.endDatePlanning : entry.startDate,
          y: i
        });
        fillLayerBackgroundColors.push(planningTaskColor);
      } else {
        // start date is earlier than planned, does not need fill
        fillLayerDataSet.push({x: entry.startDate, y: i});
        fillLayerMaskDataSet.push({x: entry.startDate, y: i});
        fillLayerBackgroundColors.push(maskColor);
      }

      if (entry.endDate !== undefined) {
        // push end date bar, if has end date
        endDateDataSet.push({x: entry.endDate, y: i});
        taskBackgroundColors.push(taskColor);
      } else {
        // add today as end date bar
        const today = new Date();
        endDateDataSet.push({x: today, y: i});
        const delta = entry.endDatePlanning.getTime() - entry.startDatePlanning.getTime();
        if (today.getTime() - entry.startDate.getTime() > delta) {
          taskBackgroundColors.push(overDueTaskColor);
        } else {
          taskBackgroundColors.push(healthProgressingTaskColor);
        }
      }
    } else {
      // push empty bars
      startDateDataSet.push({});
      endDateDataSet.push({});
      fillLayerDataSet.push({});
      fillLayerMaskDataSet.push({});
      fillLayerBackgroundColors.push(maskColor);
      taskBackgroundColors.push(maskColor);
    }
  }
  return {
    yLabels: yLabels,
    datasets: [
      {
        label: 'fillMask',
        backgroundColor: maskColor,
        data: fillLayerMaskDataSet
      },
      {
        label: 'fill',
        backgroundColor: fillLayerBackgroundColors,
        data: fillLayerDataSet,
        barPercentage: 0.8
      },
      {
        label: 'startDate',
        backgroundColor: maskColor,
        data: startDateDataSet,
        barPercentage: 0.6
      },
      {
        label: 'endDate',
        backgroundColor: taskBackgroundColors,
        data: endDateDataSet,
        barPercentage: 0.5
      },
      {
        label: 'plannedStartDate',
        backgroundColor: maskColor,
        data: startDatePlanningDataSet
      },
      {
        label: 'plannedEndDate',
        backgroundColor: planningTaskColor,
        data: endDatePlanningDataSet,
        barPercentage: 0.8
      },
    ]
  }
};
const generateChartOptions = (filteredData, timeUnit, yLabels, tickConfig) => {
  return {
    title: {
      fontSize: 25,
      position: 'top',
      padding: 15,
      display: true,
      text: '运营计划'
    },
    legend: {
      display: false,
    },
    tooltips: {
      enabled: true,
      position: 'event',
      filter: (item, data) => {
        const label = data.datasets[item.datasetIndex].label;
        const notFill = label !== 'fill';
        const notFillMask = label !== 'fillMask';
        const entry = filteredData[item.index];
        const hasStartDate = label === 'startDate' ? entry.startDate !== undefined : true;
        const hasEndDate = label === 'endDate' ? entry.endDate !== undefined : true;
        return notFill && notFillMask && hasStartDate && hasEndDate;
      },
      callbacks: {
        label: (item, data) => {
          const date = new Date(item.value);
          let label;
          switch (data.datasets[item.datasetIndex].label) {
            case 'startDate':
              label = 'Start Date';
              break;
            case 'endDate':
              label = 'End Date';
              break;
            case 'plannedStartDate':
              label = 'Planned Start Date';
              break;
            case 'plannedEndDate':
              label = 'Planned End Date';
              break;
          }
          const y = date.getFullYear();
          const m = date.getMonth() + 1;
          const d = date.getDate();
          return `${label}: ${y}/${m}/${d}`;
        },
      }
    },
    scales: {
      xAxes: [{
        type: 'time',
        gridLines: {
          display: false,
        },
        bounds: 'ticks',
        time: {
          unit: timeUnit
        },
        ticks: tickConfig
      }],
      yAxes: [{
        labels: yLabels,
        stacked: true,
        type: 'category',
      }]
    }
  }
};


export default class App extends React.Component {

  constructor(props, context) {
    super(props, context);
    this.state = {};
    this.createChart = this.createChart.bind(this);
    this.updateChart = this.updateChart.bind(this);
    this.filterData = this.filterData.bind(this);
    this.loadData = this.loadData.bind(this);
    this.selectedFilterOptionChanged = this.selectedFilterOptionChanged.bind(this);
    this.selectedQuarterOptionChanged = this.selectedQuarterOptionChanged.bind(this);
    this.selectedMonthOptionChanged = this.selectedMonthOptionChanged.bind(this);
    this.selectedYearOptionChanged = this.selectedYearOptionChanged.bind(this);
    this.chart = undefined;
  }

  componentDidMount() {
    const today = new Date();
    const min = new Date(new Date().setDate(1));
    const max = new Date(new Date(`${today.getFullYear()}/${today.getMonth() + 2}/1`).setDate(0));
    this.setState({
      tickConfig: {min: min, max: max},
      timeUnit: 'day',
      selectedMonthOption: today.getMonth() + 1,
      yearOptions: [{key: today.getFullYear(), text: today.getFullYear().toString()}],
      selectedYearOption: today.getFullYear(),
    }, this.loadData);
  }

  loadData = async () => {
    try {
      await Excel.run(async context => {
        const range = context.workbook.getSelectedRange();
        range.load("text");
        await context.sync();
        let entry;
        let data = [];
        let statusFilterOptions = [];
        let statusSet = new Set(['All']);
        let yearOptions = [];
        let yearSet = new Set();
        for (entry of range.text) {
          const entryData = {
            customerID: entry[1],
            category: entry[6],
            revenueContribution: parseFloat(entry[7]),
            startDatePlanning: new Date(entry[9]),
            endDatePlanning: new Date(entry[10]),
            startDate: entry[12] === '' ? undefined : new Date(entry[12]),
            endDate: entry[13] === '' ? undefined : new Date(entry[13]),
            status: entry[15]
          };
          data.push(entryData);
          statusSet.add(entryData.status);
          yearSet.add(entryData.startDatePlanning.getFullYear());
          yearSet.add(entryData.endDatePlanning.getFullYear());
          if (entryData.startDate !== undefined) {
            yearSet.add(entryData.startDate.getFullYear());
          }
          if (entryData.endDate !== undefined) {
            yearSet.add(entryData.endDate.getFullYear());
          }
        }
        statusSet.forEach((status) => {
          statusFilterOptions.push({key: status, text: status});
        });
        const thisYear = new Date().getFullYear();
        yearSet.add(thisYear);
        yearSet.forEach((year) => {
          yearOptions.push({key: year, text: year.toString()});
        });

        await this.setState((state) => {
          return {
            rawData: data.sort((a, b) => b.revenueContribution - a.revenueContribution),
            filterOptions: statusFilterOptions,
            yearOptions: yearOptions,
            selectedFilterOption: 'All',
            selectedYearOption: yearOptions[0].key
          }
        });
        if (this.chart === undefined) {
          this.createChart();
        } else {
          this.updateChart();
        }
      });
    } catch (error) {
      console.log(error);
    }
  };

  filterData() {
    let data = this.state.rawData;
    // filter with status
    data = data.filter((entry) => {
      if (this.state.selectedFilterOption === 'All') {
        return true;
      } else {
        return entry.status === this.state.selectedFilterOption
      }
    });
    // filter with tick bounds
    data = data.filter((entry) => {
      const min = this.state.tickConfig.min;
      const max = this.state.tickConfig.max;
      const taskNotEndTooEarly = entry.endDate !== undefined ? entry.endDate.getTime() > min.getTime() : true;
      const taskNotStartTooLate = entry.startDate !== undefined ? entry.startDate.getTime() < max.getTime() : true;
      const planNotEndTooEarly = entry.endDatePlanning.getTime() > min.getTime();
      const planNotStartTooLate = entry.startDate.getTime() < max.getTime();
      return (taskNotEndTooEarly || planNotEndTooEarly) && (taskNotStartTooLate || planNotStartTooLate);
    });
    return data;
  }

  updateChart() {
    // process data
    const filteredData = this.filterData();
    const drawConfig = processRawData(filteredData);
    // generate new chart options
    const options = generateChartOptions(
      filteredData,
      this.state.timeUnit,
      drawConfig.yLabels,
      this.state.tickConfig
    );
    this.chart.data.datasets = drawConfig.datasets;
    this.chart.options = options;
    this.chart.update();
  }

  createChart() {
    // process data
    const filteredData = this.filterData();
    const drawConfig = processRawData(filteredData);
    // config chart options
    const options = generateChartOptions(
      filteredData,
      this.state.timeUnit,
      drawConfig.yLabels,
      this.state.tickConfig
    );
    // create chart
    const ctx = document.getElementById('chart-canvas').getContext('2d');
    this.chart = new Chart(ctx, {
      type: 'horizontalBar',
      data: {
        datasets: drawConfig.datasets
      },
      options: options
    });
  }

  selectedFilterOptionChanged(event, item) {
    this.setState({selectedFilterOption: item.key}, this.updateChart);
  }

  selectedYearOptionChanged(event, item) {
    const min = new Date(this.state.tickConfig.min.setFullYear(item.key));
    const max = new Date(this.state.tickConfig.max.setFullYear(item.key));
    this.setState({
      selectedYearOption: item.key,
      tickConfig: {
        min: min,
        max: max,
      }
    }, this.updateChart);
  }

  selectedQuarterOptionChanged(event, item) {
    let min, max;
    const year = this.state.selectedYearOption;
    switch (item.key) {
      case 1:
        min = new Date(`${year}/1/1`);
        max = new Date(`${year}/3/31`);
        break;
      case 2:
        min = new Date(`${year}/4/1`);
        max = new Date(`${year}/6/30`);
        break;
      case 3:
        min = new Date(`${year}/7/1`);
        max = new Date(`${year}/9/30`);
        break;
      case 4:
        min = new Date(`${year}/10/1`);
        max = new Date(`${year}/12/31`);
        break;
    }
    this.setState({
      selectedView: 'quarter',
      selectedQuarterOption: item.key,
      selectedMonthOption: undefined,
      timeUnit: 'week',
      tickConfig: {min: min, max: max}
    }, this.updateChart);
  }

  selectedMonthOptionChanged(event, item) {
    const year = this.state.selectedYearOption;
    const min = new Date(`${year}/${item.key}/1`);
    const max = new Date(new Date(`${year}/${item.key + 1}/1`).setDate(0));
    console.log(max);
    this.setState({
      selectedView: 'month',
      selectedMonthOption: item.key,
      selectedQuarterOption: undefined,
      timeUnit: 'day',
      tickConfig: {min: min, max: max}
    }, this.updateChart);
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Stack horizontal tokens={{childrenGap: 20}} verticalAlign="end">
          <Dropdown
            label={'Filter by Status'}
            selectedKey={ this.state.selectedFilterOption ? this.state.selectedFilterOption : -1}
            onChange={this.selectedFilterOptionChanged}
            placeholder="Select Status"
            options={this.state.filterOptions}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            label={'Select Year'}
            selectedKey={this.state.selectedYearOption ? this.state.selectedYearOption : -1}
            options={this.state.yearOptions}
            onChange={this.selectedYearOptionChanged}
            placeholder={'Select Year'}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            label={'Select Quarter'}
            selectedKey={this.state.selectedQuarterOption ? this.state.selectedQuarterOption : -1}
            onChange={this.selectedQuarterOptionChanged}
            placeholder={'Select Quarter'}
            options={[
              {key: 1, text: 'Q1'},
              {key: 2, text: 'Q2'},
              {key: 3, text: 'Q3'},
              {key: 4, text: 'Q4'},
            ]}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            label={'Select Month'}
            selectedKey={this.state.selectedMonthOption ? this.state.selectedMonthOption : -1}
            onChange={this.selectedMonthOptionChanged}
            placeholder={'Select Month'}
            options={[
              {key: 1, text: 'January'},
              {key: 2, text: 'February'},
              {key: 3, text: 'March'},
              {key: 4, text: 'April'},
              {key: 5, text: 'May'},
              {key: 6, text: 'June'},
              {key: 7, text: 'July'},
              {key: 8, text: 'August'},
              {key: 9, text: 'September'},
              {key: 10, text: 'October'},
              {key: 11, text: 'November'},
              {key: 12, text: 'December'},
            ]}
            styles={{ dropdown: { width: 150 } }}
          />
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.loadData}
          >
            Reload Data
          </Button>
        </Stack>
        <canvas id={`chart-canvas`} />
      </div>
    );
  }
}
