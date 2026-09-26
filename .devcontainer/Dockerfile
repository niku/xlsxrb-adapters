FROM ruby:4.0.5

ARG USERNAME=vscode
ARG USER_UID=1000
ARG USER_GID=$USER_UID

# Create a non-root user (vscode) with passwordless sudo privileges
RUN apt-get update && apt-get install -y sudo ca-certificates curl git \
    && if [ "$(getent group ${USER_GID})" ]; then \
    echo "GID ${USER_GID} already exists."; \
    else \
    groupadd --gid ${USER_GID} ${USERNAME}; \
    fi \
    && if [ "$(getent passwd ${USER_UID})" ]; then \
    echo "UID ${USER_UID} already exists."; \
    else \
    useradd --uid ${USER_UID} --gid ${USER_GID} -m ${USERNAME} -s /bin/bash; \
    fi \
    && echo "${USERNAME} ALL=(root) NOPASSWD:ALL" > /etc/sudoers.d/${USERNAME} \
    && chmod 0440 /etc/sudoers.d/${USERNAME} \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Pre-create workspace folders and user directories with correct ownership
RUN mkdir -p /workspaces/xlsxrb-adapters /workspaces/xlsxrb /home/${USERNAME}/.local \
    && chown -R ${USERNAME}:${USERNAME} /workspaces /home/${USERNAME}

USER ${USERNAME}
WORKDIR /workspaces/xlsxrb-adapters
