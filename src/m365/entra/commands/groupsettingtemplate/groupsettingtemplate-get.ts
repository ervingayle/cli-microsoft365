import { GroupSettingTemplate } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import aadCommands from '../../aadCommands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
}

class EntraGroupSettingTemplateGetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTINGTEMPLATE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Entra group settings template';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUPSETTINGTEMPLATE_GET];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id &&
          !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.GROUPSETTINGTEMPLATE_GET, commands.GROUPSETTINGTEMPLATE_GET);

    try {
      const templates = await odata.getAllItems<GroupSettingTemplate>(`${this.resource}/v1.0/groupSettingTemplates`);

      const groupSettingTemplate: GroupSettingTemplate[] = templates.filter(t => args.options.id ? t.id === args.options.id : t.displayName === args.options.displayName);

      if (groupSettingTemplate && groupSettingTemplate.length > 0) {
        await logger.log(groupSettingTemplate.pop());
      }
      else {
        throw `Resource '${(args.options.id || args.options.displayName)}' does not exist.`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraGroupSettingTemplateGetCommand();