import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as spoGroupGetCommand from '../group/group-get';
import * as spoUserGetCommand from '../user/user-get';
import { Options as SpoGroupGetCommandOptions } from '../group/group-get';
import { Options as SpoUserGetCommandOptions } from '../user/user-get';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  force?: boolean;
}

class SpoFolderRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from the specified folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        principalId: typeof args.options.principalId !== 'undefined',
        upn: typeof args.options.upn !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folderUrl <folderUrl>'
      },
      {
        option: '--principalId [principalId]'
      },
      {
        option: '--upn [upn]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        const principalOptions: any[] = [args.options.principalId, args.options.upn, args.options.groupName];
        if (principalOptions.some(item => item !== undefined) && principalOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either principalId id, upn or groupName`;
        }

        if (principalOptions.filter(item => item !== undefined).length === 0) {
          return `Specify at least principalId id, upn or groupName`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderUrl', 'upn', 'groupName');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing role assignment from folder in site at ${args.options.webUrl}...`);
      }
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
      const requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;

      try {
        if (args.options.upn) {
          args.options.principalId = await this.getUserPrincipalId(args.options);
          await this.removeRoleAssignment(requestUrl, logger, args.options);
        }
        else if (args.options.groupName) {
          args.options.principalId = await this.getGroupPrincipalId(args.options);
          await this.removeRoleAssignment(requestUrl, logger, args.options);
        }
        else {
          await this.removeRoleAssignment(requestUrl, logger, args.options);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeRoleAssignment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove a role assignment from the folder with url '${args.options.folderUrl}'?` });

      if (result) {
        await removeRoleAssignment();
      }
    }
  }

  private async removeRoleAssignment(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/roleassignments/removeroleassignment(principalid='${options.principalId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getGroupPrincipalId(options: Options): Promise<number> {
    const groupGetCommandOptions: SpoGroupGetCommandOptions = {
      webUrl: options.webUrl,
      name: options.groupName,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(spoGroupGetCommand as Command, { options: { ...groupGetCommandOptions, _: [] } });
    const getGroupOutput = JSON.parse(output.stdout);
    return getGroupOutput.Id as number;
  }

  private async getUserPrincipalId(options: Options): Promise<number> {
    const userGetCommandOptions: SpoUserGetCommandOptions = {
      webUrl: options.webUrl,
      email: options.upn,
      id: undefined,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(spoUserGetCommand as Command, { options: { ...userGetCommandOptions, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.Id as number;
  }
}

module.exports = new SpoFolderRoleAssignmentRemoveCommand();