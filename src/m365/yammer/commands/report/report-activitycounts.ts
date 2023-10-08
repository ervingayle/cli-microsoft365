import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class YammerReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on the amount of Yammer activity in your organization by how many messages were posted, read, and liked';
  }
}

module.exports = new YammerReportActivityCountsCommand();
