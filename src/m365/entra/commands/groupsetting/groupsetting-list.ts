import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import aadCommands from '../../aadCommands';

class EntraGroupSettingListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_LIST;
  }

  public get description(): string {
    return 'Lists Entra group settings';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUPSETTING_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.GROUPSETTING_LIST, commands.GROUPSETTING_LIST);

    try {
      const groupSettings = await odata.getAllItems<GroupSetting>(`${this.resource}/v1.0/groupSettings`);
      await logger.log(groupSettings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraGroupSettingListCommand();