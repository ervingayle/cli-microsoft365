import assert from 'assert';
import child_process from 'child_process';
import fs from 'fs';
import sinon, { SinonSandbox } from 'sinon';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { formatting } from '../../../utils/formatting.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command, { SpfxDoctorCheck } from './spfx-doctor.js';

describe(commands.DOCTOR, () => {
  let log: string[];
  let sandbox: SinonSandbox;
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogSpyErr: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const packageVersionResponse = (name: string, version: string): string => {
    return `{
      "dependencies": {
        "${name}": {
          "version": "${version}"
        }
      }
    }`;
  };
  const getStatus = (status: number, message: string): string => {
    return formatting.getStatus(status, message);
  };

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogSpyErr = sinon.spy(logger, 'logToStderr');
    sinon.stub(process, 'platform').value('linux');
    (command as any).resultsObject = [];
    (command as any).output = '';
    (command as any).logger = null;
  });

  afterEach(() => {
    sinonUtil.restore([
      sandbox,
      child_process.exec,
      process.platform,
      fs.existsSync,
      fs.readFileSync,
      (command as any).getProjectVersion
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DOCTOR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes all checks for SPFx v1.11 project when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.22.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.22.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.11 project when all requirements met (debug)', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.11 generator installed locally when all requirements met (debug)', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.11 generator installed globally when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? packageVersionResponse(packageName, '1.11.0') : '{ }');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.11 generator installed locally when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.10 project when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.10 project when all requirements met (debug)', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.10 generator installed locally when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.10 generator installed locally when all requirements met (debug)', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.10 generator installed globally when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? packageVersionResponse(packageName, '1.10.0') : '{ }');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('passes all checks for SPFx v1.11 generator installed globally, SPFx version specified through args, when all requirements met', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? packageVersionResponse(packageName, '1.11.0') : '{ }');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { spfxVersion: '1.11.0' } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
  });

  it('fails with error when SPFx not found', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        case '@microsoft/generator-sharepoint':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('SharePoint Framework not found'));
    assert(loggerLogSpy.calledWith(getStatus(1, 'SharePoint Framework')), 'SharePoint Framework found');
    assert(!loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
  });

  it('fails with error when SPFx not found (debug)', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        case '@microsoft/generator-sharepoint':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        default:
          callback(new Error(`${file} ENOENT`));
          return {} as child_process.ChildProcess;
      }
    });

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('SharePoint Framework not found'));
    assert(loggerLogSpy.calledWith(getStatus(1, 'SharePoint Framework')), 'SharePoint Framework found');
    assert(!loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
  });

  it('fails with error when SPFx not found, version passed explicitly', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        case '@microsoft/generator-sharepoint':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { spfxVersion: '1.11.0' } });
    assert(loggerLogSpy.calledWith(`${getStatus(1, 'SharePoint Framework')} v1.11.0 not found`), 'SharePoint Framework found');
    assert(loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
  });

  it('passes SPO compatibility check for SPFx v1.11.0', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { env: 'spo' } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SPO')));
  });

  it('passes SPO compatibility check for SPFx v1.10.0', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { env: 'spo' } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SPO')));
  });

  it('passes SP2019 compatibility check for SPFx v1.4.1', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.4.1'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { env: 'sp2019' } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SP2019')));
  });

  it('fails validation if output does not equal text or json.', async () => {
    const actual = await command.validate({
      options: {
        output: 'md'
      }
    }, commandInfo);

    assert.notStrictEqual(actual, true);
  });

  it('fails SP2019 compatibility check for SPFx v1.5.0', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.5.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await assert.rejects(command.action(logger, { options: { env: 'sp2019' } } as any), new CommandError('SharePoint Framework v1.5.0 is not supported in SP2019'));
    assert(loggerLogSpy.calledWith(getStatus(1, 'Not supported in SP2019')));
    assert(loggerLogSpy.calledWith('- Use SharePoint Framework v1.4.1'), 'No fix provided');
  });

  it('passes SP2016 compatibility check for SPFx v1.1.0', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.1.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { env: 'sp2016' } });
    assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SP2016')));
  });

  it('fails SP2016 compatibility check for SPFx v1.2.0', async () => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.2.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await assert.rejects(command.action(logger, { options: { env: 'sp2016' } } as any), new CommandError('SharePoint Framework v1.2.0 is not supported in SP2016'));
    assert(loggerLogSpy.calledWith(getStatus(1, 'Not supported in SP2016')));
    assert(loggerLogSpy.calledWith('- Use SharePoint Framework v1.1'), 'No fix provided');
  });

  it('passes Node check when version meets single range prerequisite', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')));
  });

  it('passes Node check when version meets double range prerequisite', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.9.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'Node v8.0.0')));
  });

  it('fails Node check when version does not meet single range prerequisite', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v12.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(1, 'Node v12.0.0 found, v^10 required')));
    assert(loggerLogSpy.calledWith('- Install Node.js v10'), 'No fix provided');
  });

  it('fails Node check when version does not meet double range prerequisite', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v12.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.9.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(1, 'Node v12.0.0 found, v^8 || ^10 required')));
    assert(loggerLogSpy.calledWith('- Install Node.js v10'), 'No fix provided');
  });

  it('fails with friendly error message when npm not found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('npm not found'));
    assert(!loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
  });

  it('determines the current version from .yo-rc.json when spfxVersion not passed', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    const yoRcJson = `{
      "@microsoft/generator-sharepoint": {
        "version": "1.4.1",
        "libraryName": "spfx-141",
        "libraryId": "dd1a0a8d-e043-4ca0-b9a4-256e82a66177",
        "environment": "spo"
      }
    }`;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return yoRcJson;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { toVersion: '1.4.1', debug: true } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
  });

  it('determines the current version from .yo-rc.json when available when spfxVersion is passed', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    const yoRcJson = `{
      "@microsoft/generator-sharepoint": {
        "version": "1.4.1",
        "libraryName": "spfx-141",
        "libraryId": "dd1a0a8d-e043-4ca0-b9a4-256e82a66177",
        "environment": "spo"
      }
    }`;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return yoRcJson;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { spfxVersion: '1.4.1' } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
  });

  it('passes yo check when yo found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')));
  });

  it('fails yo check when yo not found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(1, 'yo not found')));
    assert(loggerLogSpy.calledWith('- npm i -g yo@3'), 'No fix provided');
  });

  it('fails yo check when yo not found and logs error object', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: { output: 'json' } });
    const checks: SpfxDoctorCheck[] = loggerLogSpy.lastCall.args[0];
    assert.strictEqual(checks.filter(y => y.check === 'yo')[0].passed, false);
    assert(loggerLogSpyErr.calledWith('- npm i -g yo@3'), 'No fix provided');
    assert(loggerLogSpyErr.calledWith(getStatus(1, 'yo not found')));
  });

  it('passes gulp-cli check when gulp-cli found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');

    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`), '', '');
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')));
  });

  it('fails gulp-cli check when gulp-cli not found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');

    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        default:
          callback(new Error(`${file} ENOENT`), '', '');
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(1, 'gulp-cli not found')));
    assert(loggerLogSpy.calledWith('- npm i -g gulp-cli@2'), 'No fix provided');
  });

  it('fails gulp check when gulp is found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'gulp':
          callback(null, packageVersionResponse(packageName, '4.0.0'), '');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(1, 'gulp should be removed')));
    assert(loggerLogSpy.calledWith('- npm un -g gulp'), 'No fix provided');
  });

  it('passes typescript check when typescript not found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')));
  });

  it('fails typescript check when typescript found', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'typescript':
          callback(undefined, packageVersionResponse(packageName, '3.7.5'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(1, 'typescript v3.7.5 installed in the project')));
    assert(loggerLogSpy.calledWith('- npm un typescript'), 'No fix provided');
  });

  it('returns error when used with an unsupported version of spfx', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '0.9.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(`spfx doctor doesn't support SPFx v0.9.0 at this moment`));
  });

  it('uses alternative symbols for win32', async () => {
    sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sandbox.stub(process, 'platform').value('win32');
    if (process.env.CI) {
      sandbox.stub(process.env, 'CI').value('false');
    }
    if (process.env.TERM) {
      sandbox.stub(process.env, 'TERM').value('16');
    }
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        default:
          callback({ message: `${file} ENOENT` } as any, '', '');
      }
      return {} as child_process.ChildProcess;
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
  });

  it('supports specifying environment', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '-e, --env [env]') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('passes validation when no options specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when sp2016 env specified', async () => {
    const actual = await command.validate({ options: { env: 'sp2016' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when sp2019 env specified', async () => {
    const actual = await command.validate({ options: { env: 'sp2019' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when spo env specified', async () => {
    const actual = await command.validate({ options: { env: 'spo' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when 2016 env specified', async () => {
    const actual = await command.validate({ options: { env: '2016' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when supported version of SPFx specified', async () => {
    const actual = await command.validate({ options: { spfxVersion: '1.15.2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when unsupported version of SPFx specified', async () => {
    const actual = await command.validate({ options: { spfxVersion: '1.2.3' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when supported version of SPFx prefixed with v specified', async () => {
    const actual = await command.validate({ options: { spfxVersion: 'v1.15.2' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
