import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { cli } from '../../../../cli/cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as spoGroupGetCommand from '../group/group-get';
import * as spoUserGetCommand from '../user/user-get';
const command: Command = require('./folder-roleassignment-remove');
import { settingsNames } from '../../../../settingsNames';

describe(commands.FOLDER_ROLEASSIGNMENT_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.executeCommandWithOutput,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoUserGetCommand) {
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
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoUserGetCommand) {
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

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
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
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
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
    assert(promptIssued);
  });

  it('removes role assignment when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
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
