FROM dockdock/libreoffice73pyenvpoetry:lo-7.5.3-RELEASE
ENV PYTHONPATH=/workspaces/libreoffice-modern-code-editing-py
RUN poetry config virtualenvs.in-project true \
    && git config --global --add safe.directory '*'