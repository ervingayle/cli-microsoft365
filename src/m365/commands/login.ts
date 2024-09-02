import * as fs from 'fs';
import auth, { AuthType, CloudType } from '../../Auth';
import Command, { CommandError } from '../../Command';
import { Logger } from '../../cli/Logger';
import { cli } from '../../cli/cli';
import { settingsNames } from '../../settingsNames';
import commands from './commands';
import GlobalOptions from '../../GlobalOptions';
import { misc } from '../../utils/misc';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  authType?: string;
  cloud?: string;
  userName?: string;
  password?: string;
  certificateFile?: string;
  certificateBase64Encoded?: string;
  thumbprint?: string;
  appId?: string;
  tenant?: string;
  secret?: string;
}

class LoginCommand extends Command {
  private static allowedAuthTypes: string[] = ['certificate', 'deviceCode', 'password', 'identity', 'browser', 'secret'];

  public get name(): string {
    return commands.LOGIN;
  }

  public get description(): string {
    return 'Log in to Microsoft 365';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        authType: args.options.authType || 'deviceCode',
        cloud: args.options.cloud ?? CloudType.Public
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --authType [authType]',
        autocomplete: LoginCommand.allowedAuthTypes
      },
      {
        option: '-u, --userName [userName]'
      },
      {
        option: '-p, --password [password]'
      },
      {
        option: '-c, --certificateFile [certificateFile]'
      },
      {
        option: '--certificateBase64Encoded [certificateBase64Encoded]'
      },
      {
        option: '--thumbprint [thumbprint]'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '--tenant [tenant]'
      },
      {
        option: '-s, --secret [secret]'
      },
      {
        option: '--cloud [cloud]',
        autocomplete: misc.getEnums(CloudType)
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.authType === 'password') {
          if (!args.options.userName) {
            return 'Required option userName missing';
          }

          if (!args.options.password) {
            return 'Required option password missing';
          }
        }

        if (args.options.authType === 'certificate') {
          if (args.options.certificateFile && args.options.certificateBase64Encoded) {
            return 'Specify either certificateFile or certificateBase64Encoded, but not both.';
          }

          if (!args.options.certificateFile && !args.options.certificateBase64Encoded) {
            return 'Specify either certificateFile or certificateBase64Encoded';
          }

          if (args.options.certificateFile) {
            if (!fs.existsSync(args.options.certificateFile)) {
              return `File '${args.options.certificateFile}' does not exist`;
            }
          }
        }

        if (args.options.authType &&
          LoginCommand.allowedAuthTypes.indexOf(args.options.authType) < 0) {
          return `'${args.options.authType}' is not a valid authentication type. Allowed authentication types are ${LoginCommand.allowedAuthTypes.join(', ')}`;
        }

        if (args.options.authType === 'secret') {
          if (!args.options.secret) {
            return 'Required option secret missing';
          }
        }

        if (args.options.cloud &&
          typeof CloudType[args.options.cloud as keyof typeof CloudType] === 'undefined') {
          return `${args.options.cloud} is not a valid value for cloud. Valid options are ${misc.getEnums(CloudType).join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    // disconnect before re-connecting
    if (this.debug) {
      await logger.logToStderr(`Logging out from Microsoft 365...`);
    }

    const deactivate: () => void = (): void => auth.connection.deactivate();

    const getCertificate: (options: Options) => string | undefined = (options): string | undefined => {
      // command args take precedence over settings
      if (options.certificateFile) {
        return fs.readFileSync(options.certificateFile).toString('base64');
      }
      if (options.certificateBase64Encoded) {
        return options.certificateBase64Encoded;
      }
      return cli.getConfig().get(settingsNames.clientCertificateFile) ||
        cli.getConfig().get(settingsNames.clientCertificateBase64Encoded);
    };

    const login: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Signing in to Microsoft 365...`);
      }

      const authType = args.options.authType || cli.getSettingWithDefaultValue<string>(settingsNames.authType, 'deviceCode');
      auth.connection.appId = args.options.appId || cli.getClientId();
      auth.connection.tenant = args.options.tenant || cli.getTenant();
      auth.connection.name = args.options.connectionName;

      switch (authType) {
        case 'password':
          auth.connection.authType = AuthType.Password;
          auth.connection.userName = args.options.userName;
          auth.connection.password = args.options.password;
          break;
        case 'certificate':
          auth.connection.authType = AuthType.Certificate;
          auth.connection.certificate = getCertificate(args.options);
          auth.connection.thumbprint = args.options.thumbprint;
          auth.connection.password = args.options.password || cli.getConfig().get(settingsNames.clientCertificatePassword);
          break;
        case 'identity':
          auth.connection.authType = AuthType.Identity;
          auth.connection.userName = args.options.userName;
          break;
        case 'browser':
          auth.connection.authType = AuthType.Browser;
          break;
        case 'secret':
          auth.connection.authType = AuthType.Secret;
          auth.connection.secret = args.options.secret || cli.getConfig().get(settingsNames.clientSecret);
          break;
      }

      if (args.options.cloud) {
        auth.connection.cloudType = CloudType[args.options.cloud as keyof typeof CloudType];
      }
      else {
        auth.connection.cloudType = CloudType.Public;
      }

      try {
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
        auth.connection.active = true;
      }
      catch (error: any) {
        if (this.debug) {
          await logger.logToStderr('Error:');
          await logger.logToStderr(error);
          await logger.logToStderr('');
        }

        throw new CommandError(error.message);
      }

      const details = auth.getConnectionDetails(auth.connection);

      if (this.debug) {
        (details as any).accessToken = JSON.stringify(auth.connection.accessTokens, null, 2);
      }

      await logger.log(details);
    };

    deactivate();
    await login();
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

module.exports = new LoginCommand();