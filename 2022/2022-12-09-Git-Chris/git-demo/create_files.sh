for a in $@; do
  file=${a}_file.txt
  echo "Created ${a}" | tee $file
  git add $file
  git commit $file -m"Created ${file}"
done
