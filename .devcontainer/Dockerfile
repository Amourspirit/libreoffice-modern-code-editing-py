FROM dockdock/libreoffice73pyenvpoetry:lo-7.5.3-RELEASE
ARG HOME_DIR=/workspaces/libreoffice-modern-code-editing-py
ENV PYTHONPATH=$HOME_DIR
RUN poetry config virtualenvs.in-project true \
    && git config --global --add safe.directory '*'

# RUN poetry install

# ENV VIRTUAL_ENV=$HOME_DIR

# RUN poetry env use "$HOME_DIR/.venv/bin/python3"
# CMD [ "source", "/workspaces/libreoffice-modern-code-editing-py/.venv/bin/activate" ]