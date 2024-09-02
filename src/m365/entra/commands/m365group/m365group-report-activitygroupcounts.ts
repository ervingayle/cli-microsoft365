import PeriodBasedReport from '../../../base/PeriodBasedReport';
import aadCommands from '../../aadCommands';
import commands from '../../commands';

class M365GroupReportActivityGroupCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.M365GROUP_REPORT_ACTIVITYGROUPCOUNTS;
  }

  public get description(): string {
    return 'Get the daily total number of groups and how many of them were active based on email conversations, Viva Engage posts, and SharePoint file activities';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_REPORT_ACTIVITYGROUPCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityGroupCounts';
  }
}

module.exports = new M365GroupReportActivityGroupCountsCommand();