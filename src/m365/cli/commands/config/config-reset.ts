import { cli } from "../../../../cli/cli";
import { Logger } from "../../../../cli/Logger";
import GlobalOptions from "../../../../GlobalOptions";
import { settingsNames } from "../../../../settingsNames";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key?: string;
}

class CliConfigResetCommand extends AnonymousCommand {
  private static readonly optionNames: string[] = Object.getOwnPropertyNames(settingsNames);

  public get name(): string {
    return commands.CONFIG_RESET;
  }

  public get description(): string {
    return 'Resets the specified CLI configuration option to its default value';
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
        key: args.options.key
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-k, --key [key]',
        autocomplete: CliConfigResetCommand.optionNames
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.key) {
          if (CliConfigResetCommand.optionNames.indexOf(args.options.key) < 0) {
            return `${args.options.key} is not a valid setting. Allowed values: ${CliConfigResetCommand.optionNames.join(', ')}`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.key) {
      cli.getConfig().delete(args.options.key);
    }
    else {
      cli.getConfig().clear();
    }
  }
}

module.exports = new CliConfigResetCommand();
