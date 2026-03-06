import paramiko

from app.core.logger import logger, log_function


class SSH_setup:
    def __init__(self):
        self.is_connect = False
        self.ssh = None
        #self.host = "192.168.1.2"
        self.host = "10.119.9.225"
        self.port = 22
        self.username = "robot"
        self.password = "robot"
        self.script_path = "/home/robot/Manufacturing_test/aipc_beta/test.py ecat"
        self.timeout = 10  # seconds
        #self.config = self.load_config()
        self.logger2 = logger.getChild('SSH_setup')


    @log_function
    def Connect_RPI(self, host=None, port=None, username=None, password=None):
        try:
            # Use provided parameters or fall back to instance variables
            host = host or self.host
            port = port or self.port
            username = username or self.username
            password = password or self.password

            self.ssh = paramiko.SSHClient()
            self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            self.ssh.connect(
                host,
                port,
                username,
                password,
                timeout=self.timeout
            )
            self.is_connect = True

            # logger.info("SSH connection established successfully")
            self.logger2.info(f"SSH connection established successfully",
                              extra={'func_name': 'Connect_RPI'})

            return True, "Connected successfully"
        except paramiko.AuthenticationException:
            error_msg = "Authentication failed, please verify your credentials"
            self.logger2.error(f"{error_msg}", exc_info=True,
                               extra={'func_name': 'Connect_RPI'})
            return False, error_msg
        except paramiko.SSHException as e:
            error_msg = f"SSH error: {str(e)}"
            self.logger2.error(f"{error_msg}", exc_info=True,
                               extra={'func_name': 'Connect_RPI'})
            return False, error_msg
        except Exception as e:
            error_msg = f"Connection error: {str(e)}"
            self.logger2.error(f"{error_msg}", exc_info=True,
                               extra={'func_name': 'Connect_RPI'})
            return False, error_msg

    def SSH_com(self, command, script_path=None):
        if not self.is_connect or not self.ssh:
            return "", "Not connected to SSH"


        script_path = script_path or self.script_path

        try:
            stdin, stdout, stderr = self.ssh.exec_command(f'sudo python3 {script_path} {command}')
            stdout_data = stdout.read().decode()
            stderr_data = stderr.read().decode()
            return stdout_data, stderr_data
        except Exception as e:
            return "", f"Command execution failed: {str(e)}"

    def SSH_com_stream(self, script_path, command):
        if not self.is_connect:
            raise Exception("SSH connection not established")

        # Run the Python script
        stdin, stdout, stderr = self.ssh.exec_command(f'sudo python3 {script_path} {command}', get_pty=True)

        # Continuously read the output and error
        while True:
            line = stdout.readline()
            if not line:
                break
            yield line.strip()

    @log_function
    def SSH_disconnect(self):
        try:
            if self.ssh:
                self.ssh.close()
                # logger.info("SSH connection closed")
                self.logger2.info(f"SSH connection closed",
                                  extra={'func_name': 'SSH_disconnect'})

        except Exception as e:
            # logger.error(f"Error disconnecting SSH: {str(e)}")
            self.logger2.error(f"Error disconnecting SSH: {str(e)}", exc_info=True,
                               extra={'func_name': 'SSH_disconnect'})
        finally:
            self.is_connect = False
