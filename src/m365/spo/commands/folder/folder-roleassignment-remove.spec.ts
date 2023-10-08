import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import commands from '../../commands';
import * as SpoUserGetCommand from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
const command: Command = require('./folder-roleassignment-remove');
import { settingsNames } from '../../../../settingsNames';

describe(commands.FOLDER_ROLEASSIGNMENT_REMOVE, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptOptions: any;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    requests = [];
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.executeCommandWithOutput,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ROLEASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if principalId and upn are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, upn: 'someaccount@tenant.onmicrosoft.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, groupName: 'someGroup' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if upn and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', upn: 'someaccount@tenant.onmicrosoft.com', groupName: 'someGroup' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither upn nor principalId or groupName is specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if folderUrl is not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', upn: 'someaccount@tenant.onmicrosoft.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('remove role assignment from folder by folderUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        principalId: 11,
        force: true
      }
    });
  });

  it('remove role assignment from folder and get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
        };
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no user found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('remove role assignment from folder and get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        force: true
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('aborts removing role assignment when prompt not confirmed', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      }
    });

    assert(requests.length === 0);
  });

  it('prompts before removing role assignment when confirmation argument not passed', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('removes role assignment when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      }
    });
  });
});
