import { Logger } from '../../../../cli/Logger';
import PeriodBasedReport, { CommandArgs } from '../../../base/PeriodBasedReport';
import commands from '../../commands';
import yammerCommands from './yammerCommands';

class VivaEngageReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageDistributionUserCounts';
  }

  public get description(): string {
    return 'Gets the number of users by device type';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

module.exports = new VivaEngageReportDeviceUsageDistributionUserCountsCommand();

