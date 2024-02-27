const useOnOfficeItemReady = () => {
  const [emailItem, setEmailItem] = useState(null);
  const [officeIsReady, setOfficeIsReady] = useState(null);
  //#region Handle Office On Ready & set Default values
  const officeOnReadyCallback = useCallback(() => {
    if (officeIsReady) return;
    if (!Office.context.mailbox) return;
    setOfficeIsReady(true);
    setEmailItem(Office.context.mailbox.item);
  }, [officeIsReady]);
  useEffect(() => {
    Office.onReady(officeOnReadyCallback);
  }, [officeOnReadyCallback]);
  useEffect(() => {
    if (!officeIsReady) return;
  }, [officeIsReady]);
  //#endregion

  return {
    isReady: officeIsReady,
    emailItem,
  };
};
