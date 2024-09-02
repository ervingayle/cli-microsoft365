import { GroupSettingTemplate } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import aadCommands from '../../aadCommands';

class EntraGroupSettingTemplateListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTINGTEMPLATE_LIST;
  }

  public get description(): string {
    return 'Lists Entra group settings templates';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUPSETTINGTEMPLATE_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.GROUPSETTINGTEMPLATE_LIST, commands.GROUPSETTINGTEMPLATE_LIST);

    try {
      const templates = await odata.getAllItems<GroupSettingTemplate>(`${this.resource}/v1.0/groupSettingTemplates`);
      await logger.log(templates);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraGroupSettingTemplateListCommand();