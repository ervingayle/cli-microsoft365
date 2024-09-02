import PeriodBasedReport from '../../../base/PeriodBasedReport';
import aadCommands from '../../aadCommands';
import commands from '../../commands';

class M365GroupReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.M365GROUP_REPORT_ACTIVITYCOUNTS;
  }

  public get description(): string {
    return 'Get the number of group activities across group workloads';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_REPORT_ACTIVITYCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityCounts';
  }
}

module.exports = new M365GroupReportActivityCountsCommand();