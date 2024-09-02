import { Logger } from '../../../../cli/Logger';
import DateAndPeriodBasedReport, { CommandArgs } from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';
import yammerCommands from './yammerCommands';

class VivaEngageReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_DEVICEUSAGEUSERDETAIL];
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about Viva Engage device usage by user';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

module.exports = new VivaEngageReportDeviceUsageUserDetailCommand();